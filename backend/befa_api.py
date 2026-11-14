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
