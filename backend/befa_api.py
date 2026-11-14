"""
Routes HTTP du Business Plan BEFA (phase 2).

Lot 2 : POST /api/befa/liasse - import d'une liasse SYSCOHADA (PDF).
Le resultat est un brouillon a VALIDER par le consultant : chiffres
condenses, tableaux d'analyse, alertes de controle, note de lecture de l'IA.
"""
import base64
from dataclasses import asdict
from typing import List, Optional

from fastapi import APIRouter, File, HTTPException, UploadFile

from befa import graphiques, trends
from befa.historique import AnalyseHistorique
from befa.ia import IaError
from befa.liasse import convertir, lire_liasse

router = APIRouter(prefix="/api/befa", tags=["BEFA"])

TAILLE_MAX = 30 * 1024 * 1024  # limite de l'API Claude : 32 Mo par requete


def _lignes(lignes):
    return [{"libelle": l.libelle, "valeurs": l.valeurs, "style": l.style} for l in lignes]


def _dict(obj) -> Optional[dict]:
    return asdict(obj) if obj is not None else None


@router.post("/liasse")
async def importer_liasse(fichier: UploadFile = File(...)):
    contenu = await fichier.read()
    if not contenu.startswith(b"%PDF"):
        raise HTTPException(status_code=422, detail="Le fichier doit être un PDF (liasse fiscale SYSCOHADA).")
    if len(contenu) > TAILLE_MAX:
        raise HTTPException(status_code=413, detail="Fichier trop volumineux (30 Mo maximum).")
    try:
        lue = lire_liasse(contenu, fichier.filename or "liasse.pdf")
    except IaError as e:
        raise HTTPException(status_code=502, detail=str(e))

    imp = convertir(lue)
    analyse = AnalyseHistorique(imp.bilan_n, imp.bilan_n1, imp.cr_n, imp.cr_n1)
    return {
        "fichier": fichier.filename,
        "pages_lues": imp.pages,
        "entreprise": imp.entreprise,
        "bilan_n": _dict(imp.bilan_n),
        "bilan_n1": _dict(imp.bilan_n1),
        "compte_resultat_n": _dict(imp.cr_n),
        "compte_resultat_n1": _dict(imp.cr_n1),
        "alertes": imp.alertes,
        "controles_ok": not imp.alertes,
        "note_ia": imp.remarques,  # observation libre de l'IA, non verifiee
        "apercu": {
            "colonnes": analyse.colonnes,
            "bilan": _lignes(analyse.lignes_bilan()),
            "structure": _lignes(analyse.lignes_structure()),
            "compte_resultat": _lignes(analyse.lignes_compte_resultat()),
            "ratios_n": imp.bilan_n.ratios(),
        },
        "lecture_brute": lue.vers_json(),
    }


def _image(png: bytes) -> str:
    return "data:image/png;base64," + base64.b64encode(png).decode("ascii")


@router.post("/trends")
async def importer_trends(fichiers: List[UploadFile] = File(...)):
    """Lot 3 : analyse des exports CSV de Google Trends (un ou plusieurs fichiers)."""
    lus = []
    for f in fichiers:
        contenu = await f.read()
        if len(contenu) > 5 * 1024 * 1024:
            raise HTTPException(status_code=413, detail=f"{f.filename} : fichier trop volumineux (5 Mo maximum).")
        lus.append((f.filename or "trends.csv", contenu))
    try:
        donnees = trends.lire_fichiers(lus)
    except trends.TrendsError as e:
        raise HTTPException(status_code=422, detail=str(e))

    series, images = [], {}
    ev = donnees.evolution
    if ev:
        analyses = [trends.analyser_serie(s) for s in ev.series]
        for a in analyses:
            d = asdict(a)
            for cle in ("debut", "fin", "pic_date"):
                d[cle] = d[cle].isoformat()
            series.append(d)
        images["evolution"] = _image(graphiques.evolution_interet([(s.terme, s.points) for s in ev.series], ev.granularite))
        principale = next((a for a in analyses if a.saisonnalite), None)
        if principale:
            images["saisonnalite"] = _image(graphiques.saisonnalite(principale.saisonnalite))
    if donnees.regions and donnees.regions[0].regions:
        top = donnees.regions[0].regions[:10]
        images["regions"] = _image(graphiques.barres_horizontales([n for n, _ in top], [v for _, v in top]))

    return {
        "fichiers": donnees.fichiers,
        "categorie": ev.categorie if ev else "",
        "granularite": ev.granularite if ev else "",
        "series": series,
        "regions": [{"terme": r.terme, "regions": r.regions} for r in donnees.regions],
        "associes": [asdict(a) for a in donnees.associes],
        "faits": trends.faits(donnees),
        "source": trends.source(donnees),
        "avertissement": trends.AVERTISSEMENT,
        "graphiques": images,
    }


@router.post("/redaction")
def rediger_dossier(dossier: dict, recherche_web: bool = True):
    """Lot 4 : redaction IA des sections narratives + commentaires financiers (programme) + controles.

    Corps : le dossier BEFA au format JSON (cf. befa.dossier.DossierBEFA, ex. data/befa_exemple_coreal.json).
    """
    from befa.dossier import DossierBEFA
    from befa.projection import FinanceInputError, calculer_projection
    from befa.redaction import rediger

    try:
        d = DossierBEFA.from_dict(dossier)
    except (KeyError, TypeError) as e:
        raise HTTPException(status_code=422, detail=f"Dossier incomplet ou mal formé : {e}")
    erreurs = d.validate()
    if erreurs:
        raise HTTPException(status_code=422, detail="\n".join(erreurs))
    try:
        proj = calculer_projection(d.projection) if d.projection else None
    except FinanceInputError as e:
        raise HTTPException(status_code=422, detail=str(e))
    try:
        red = rediger(d, proj, recherche_web=recherche_web)
    except IaError as e:
        raise HTTPException(status_code=502, detail=str(e))
    return red.vers_json()


# ----------------------------------------------------------------------
# Lot 5 : generation du Business Plan (Word + PDF)
# ----------------------------------------------------------------------
import json as _json  # noqa: E402
import os as _os  # noqa: E402
import tempfile as _tempfile  # noqa: E402
import uuid as _uuid  # noqa: E402

from fastapi import Form  # noqa: E402
from fastapi.responses import FileResponse  # noqa: E402

_DOCUMENTS: dict = {}
_SORTIES = _os.path.join(_tempfile.gettempdir(), "befa_business_plans")


def _nom_fichier(texte: str) -> str:
    sur = "".join(c if c.isalnum() or c in "-_" else "_" for c in texte).strip("_")
    return sur[:60] or "business_plan"


@router.post("/generer")
async def generer_business_plan(
    dossier: str = Form(...),
    redaction: Optional[str] = Form(None),
    fichiers_trends: List[UploadFile] = File(default=[]),
):
    """Genere le Business Plan BEFA. `dossier` et `redaction` (facultative, issue de /redaction et
    eventuellement corrigee) sont des chaines JSON ; les CSV Google Trends sont facultatifs."""
    from befa.document import generer
    from befa.dossier import DossierBEFA
    from befa.pdf import PdfError
    from befa.projection import FinanceInputError
    from befa.redaction import redaction_depuis_json

    try:
        d = DossierBEFA.from_dict(_json.loads(dossier))
        red = redaction_depuis_json(_json.loads(redaction)) if redaction else None
    except (ValueError, KeyError, TypeError) as e:
        raise HTTPException(status_code=422, detail=f"Données mal formées : {e}")
    erreurs = d.validate()
    if erreurs:
        raise HTTPException(status_code=422, detail="\n".join(erreurs))
    donnees_trends = None
    lus = [(f.filename or "trends.csv", await f.read()) for f in fichiers_trends if f.filename]
    if lus:
        try:
            donnees_trends = trends.lire_fichiers(lus)
        except trends.TrendsError as e:
            raise HTTPException(status_code=422, detail=str(e))
        if not d.trends_faits:
            d.trends_faits, d.trends_source = trends.faits(donnees_trends), trends.source(donnees_trends)

    job = _uuid.uuid4().hex
    prefixe = "Business_Plan" if d.variante == "nouveau_projet" else "Dossier_credit"
    nom = f"{prefixe}_{_nom_fichier(d.entreprise.sigle or d.entreprise.raison_sociale)}"
    try:
        res = generer(d, red, donnees_trends, _os.path.join(_SORTIES, job), nom)
    except FinanceInputError as e:
        raise HTTPException(status_code=422, detail=str(e))
    except PdfError as e:
        raise HTTPException(status_code=500, detail=str(e))
    _DOCUMENTS[job] = {"pdf": res.pdf, "docx": res.docx}
    return {
        "job_id": job,
        "pages": res.pages,
        "sommaire_complet": res.sommaire_complet,
        "alertes": res.alertes + (red.alertes if red else []),
        "pdf": f"/api/befa/telecharger/{job}/pdf",
        "docx": f"/api/befa/telecharger/{job}/docx",
    }


@router.get("/telecharger/{job_id}/{format_}")
def telecharger(job_id: str, format_: str):
    doc = _DOCUMENTS.get(job_id)
    if not doc or format_ not in ("pdf", "docx") or not _os.path.exists(doc[format_]):
        raise HTTPException(status_code=404, detail="Document introuvable ou expiré : relancez la génération.")
    types = {"pdf": "application/pdf",
             "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"}
    return FileResponse(doc[format_], media_type=types[format_], filename=_os.path.basename(doc[format_]))


# ----------------------------------------------------------------------
# Lot 6 : services du formulaire web
# ----------------------------------------------------------------------
_EXEMPLES = {"coreal": "befa_exemple_coreal.json", "oivoire": "befa_exemple_oivoire.json"}
_DATA = _os.path.join(_os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))), "data")


@router.get("/exemples/{nom}")
def exemple(nom: str):
    """Dossiers d'exemple (demonstration, formation des consultants)."""
    if nom not in _EXEMPLES:
        raise HTTPException(status_code=404, detail="Exemple inconnu.")
    with open(_os.path.join(_DATA, _EXEMPLES[nom]), encoding="utf-8") as f:
        return _json.load(f)


@router.post("/projection")
def apercu_projection(corps: dict):
    """Apercu instantane des etats previsionnels pendant la saisie.

    Corps : {"projection": {...}, "nom": "COREAL FOOD"}. Repond toujours 200 : en cas de saisie
    incomplete, `erreurs` decrit ce qui manque (affiche sans bloquer la saisie)."""
    from befa import commentaires as C
    from befa.projection import FinanceInputError, ProjectionInput, calculer_projection

    nom = corps.get("nom") or "L'entreprise"
    try:
        proj = calculer_projection(ProjectionInput.from_dict(corps.get("projection") or {}))
    except FinanceInputError as e:
        return {"erreurs": [l.lstrip("- ") for l in str(e).splitlines() if l.strip() and not l.endswith(":")]}
    except (TypeError, KeyError, ValueError) as e:
        return {"erreurs": [f"Saisie incomplète : {e}"]}
    return {
        "erreurs": [],
        "annees": proj.compte.annees,
        "compte": _lignes(proj.compte.lignes()),
        "plan": {"colonnes": proj.plan.colonnes, "lignes": _lignes(proj.plan.lignes())},
        "couverture": _lignes(proj.couverture.lignes() + proj.couverture.lignes_service_dette()),
        "synthese": {
            "cout_investissements": proj.cout_total_investissements,
            "apports": proj.total_apports,
            "prets": proj.total_prets,
        },
        "commentaires": {
            "financement": C.financement(proj, nom),
            "compte_exploitation": C.compte_exploitation(proj, nom),
            "plan_financement": C.plan_financement(proj, nom),
            "couverture_dette": C.couverture_dette(proj),
        },
    }
