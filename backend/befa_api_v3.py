"""
Routes HTTP de l'evolution « constitution intelligente du dossier de financement » :

  Referentiels    GET  /api/befa/referentiels                 secteurs, banques, modeles de documents
  Administration  GET/POST /api/befa/admin/{banques|modeles}  liste / ajout ou modification
                  DELETE   /api/befa/admin/{nom}/{id}          suppression
                  POST     /api/befa/admin/{nom}/reinitialiser liste d'origine
  Entreprise      POST /api/befa/identite                     RCCM (obligatoire) + statuts -> identite
                  POST /api/befa/partenaires                  Excel clients / fournisseurs
                  GET  /api/befa/partenaires/modele           modele Excel vierge
  Comptabilite    POST /api/befa/liasses                      1 a 3 liasses (PDF / Excel) -> 3 exercices
                  POST /api/befa/bilan-ouverture              proposition de bilan d'ouverture (startup)
  Finances        POST /api/befa/tableau                      compte d'exploitation / plan de tresorerie
  Courriers       POST /api/befa/courriers/preparer           texte modifiable, un courrier par banque
                  POST /api/befa/courriers/generer            Word + PDF + ZIP
  Documents       POST /api/befa/documents/preparer           redaction IA d'un document de la liste
                  POST /api/befa/documents/generer            Word + PDF
  Fichiers        GET  /api/befa/fichier/{job}/{nom}          telechargement des fichiers generes
  Dossiers        GET/PUT/DELETE /api/befa/dossiers[/{id}]    sauvegarde serveur, versions, sources

Toute route d'import accepte un champ `dossier_id` : le document source et le
resultat de l'extraction sont alors conserves (tracabilite).
"""
import json
import os
import tempfile
import uuid
from concurrent.futures import ThreadPoolExecutor
from dataclasses import asdict
from typing import Any, Dict, List, Optional

from fastapi import APIRouter, Body, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, Response

from befa import referentiels, secteurs, stockage
from befa.extraction import ExtractionError, extraire_identite, extraire_tableau
from befa.ia import IaError

router = APIRouter(prefix="/api/befa", tags=["BEFA - constitution du dossier"])

_FICHIERS: Dict[str, Dict[str, str]] = {}  # job -> {nom de fichier: chemin}
_SORTIES = os.path.join(tempfile.gettempdir(), "befa_documents")
TYPES = {"pdf": "application/pdf", "zip": "application/zip",
         "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
         "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}


def _erreur(e: Exception, code: int = 422):
    raise HTTPException(status_code=code, detail=str(e))


def _publier(chemins: List[str]) -> Dict[str, str]:
    """Rend des fichiers telechargeables : retourne {nom: url}."""
    job = uuid.uuid4().hex
    _FICHIERS[job] = {os.path.basename(c): c for c in chemins if c}
    return {nom: f"/api/befa/fichier/{job}/{nom}" for nom in _FICHIERS[job]}


def _dossier(corps: Dict[str, Any]):
    from befa.dossier import DossierBEFA
    from befa.projection import FinanceInputError, calculer_projection

    try:
        d = DossierBEFA.from_dict(corps)
    except (KeyError, TypeError, ValueError) as e:
        raise HTTPException(status_code=422, detail=f"Dossier incomplet ou mal formé : {e}")
    proj = None
    if d.projection and d.projection.produits:
        try:
            proj = calculer_projection(d.projection)
        except (FinanceInputError, TypeError, ValueError):
            proj = None  # donnees financieres incompletes : les documents sont rediges sans les chiffres
    return d, proj


async def _lire(f: UploadFile, max_mo: int = 30) -> bytes:
    contenu = await f.read()
    if len(contenu) > max_mo * 1024 * 1024:
        raise HTTPException(status_code=413, detail=f"{f.filename} : fichier trop volumineux ({max_mo} Mo maximum).")
    if not contenu:
        raise HTTPException(status_code=422, detail=f"{f.filename} : fichier vide.")
    return contenu


@router.get("/fichier/{job}/{nom}")
def fichier(job: str, nom: str):
    chemin = _FICHIERS.get(job, {}).get(nom)
    if not chemin or not os.path.exists(chemin):
        raise HTTPException(status_code=404, detail="Fichier introuvable ou expiré : relancez la génération.")
    ext = nom.rsplit(".", 1)[-1].lower()
    return FileResponse(chemin, media_type=TYPES.get(ext, "application/octet-stream"), filename=nom)


# ----------------------------------------------------------------------
# Referentiels et administration
# ----------------------------------------------------------------------
@router.get("/referentiels")
def lister_referentiels():
    return {"secteurs": [{"code": s["code"], "libelle": s["libelle"], "sous_secteurs": s["sous_secteurs"]}
                         for s in secteurs.SECTEURS],
            "banques": referentiels.lister("banques", actifs_seulement=True),
            "modeles": referentiels.lister("modeles", actifs_seulement=True)}


@router.get("/admin/{nom}")
def admin_lister(nom: str):
    try:
        return referentiels.lister(nom)
    except referentiels.ReferentielError as e:
        _erreur(e, 404)


@router.post("/admin/{nom}")
def admin_enregistrer(nom: str, element: Dict[str, Any] = Body(...)):
    try:
        return referentiels.enregistrer(nom, element)
    except referentiels.ReferentielError as e:
        _erreur(e)


@router.delete("/admin/{nom}/{ident}")
def admin_supprimer(nom: str, ident: str):
    try:
        referentiels.supprimer(nom, ident)
    except referentiels.ReferentielError as e:
        _erreur(e, 404)
    return {"supprime": ident}


@router.post("/admin/{nom}/reinitialiser")
def admin_reinitialiser(nom: str):
    try:
        return referentiels.reinitialiser(nom)
    except referentiels.ReferentielError as e:
        _erreur(e, 404)


# ----------------------------------------------------------------------
# Page Entreprise : RCCM, statuts, clients et fournisseurs
# ----------------------------------------------------------------------
@router.post("/identite")
async def importer_identite(rccm: UploadFile = File(...), statuts: Optional[UploadFile] = File(None),
                            dossier_id: str = Form("")):
    fichiers = [(rccm.filename or "rccm.pdf", await _lire(rccm), "rccm")]
    if statuts is not None and statuts.filename:
        fichiers.append((statuts.filename, await _lire(statuts), "statuts"))
    try:
        res = extraire_identite(fichiers)
    except ExtractionError as e:
        _erreur(e)
    except IaError as e:
        _erreur(e, 502)
    for nom, octets, typ in fichiers:
        stockage.conserver_source(dossier_id or None, nom, octets, typ, res)
    return res


@router.post("/partenaires")
async def importer_partenaires(fichier: UploadFile = File(...), type: str = Form("auto"),
                               total_reference: Optional[float] = Form(None), dossier_id: str = Form("")):
    from befa import partenaires

    contenu = await _lire(fichier, 10)
    try:
        res = partenaires.importer(fichier.filename or "partenaires.xlsx", contenu, type, total_reference)
    except ExtractionError as e:
        _erreur(e)
    stockage.conserver_source(dossier_id or None, fichier.filename or "partenaires.xlsx", contenu, "partenaires", res)
    return res


@router.get("/partenaires/modele")
def modele_partenaires():
    from befa import partenaires

    return Response(partenaires.modele_excel(), media_type=TYPES["xlsx"],
                    headers={"Content-Disposition": 'attachment; filename="Modele_clients_fournisseurs_BEFA.xlsx"'})


# ----------------------------------------------------------------------
# Documents comptables
# ----------------------------------------------------------------------
@router.post("/liasses")
async def importer_liasses(fichiers: List[UploadFile] = File(...), dossier_id: str = Form("")):
    """Jusqu'a 3 liasses (PDF, Excel) lues en parallele ; exercices fusionnes (3 plus recents)."""
    from befa.historique import AnalyseHistorique
    from befa.liasse import convertir, fusionner, lire_liasse_fichier

    if not 1 <= len(fichiers) <= 3:
        raise HTTPException(status_code=422, detail="Importez de 1 à 3 liasses (les trois derniers bilans).")
    lus = [(f.filename or f"liasse_{i + 1}.pdf", await _lire(f)) for i, f in enumerate(fichiers)]

    def lire(x):
        nom, contenu = x
        try:
            return nom, convertir(lire_liasse_fichier(contenu, nom)), None
        except (IaError, ExtractionError) as e:
            return nom, None, str(e)

    with ThreadPoolExecutor(max_workers=3) as pool:
        resultats = list(pool.map(lire, lus))
    imports = [imp for _, imp, err in resultats if imp is not None]
    if not imports:
        raise HTTPException(status_code=502, detail="\n".join(f"{n} : {err}" for n, _, err in resultats))
    fus = fusionner(imports)
    for (nom, contenu), (_, imp, err) in zip(lus, resultats):
        stockage.conserver_source(dossier_id or None, nom, contenu, "liasse",
                                  {"erreur": err} if imp is None else {"alertes": imp.alertes,
                                                                       "entreprise": imp.entreprise})
    b, c = fus["bilans"] + [None] * 3, fus["crs"] + [None] * 3
    analyse = AnalyseHistorique(b[0], b[1], c[0], c[1])
    alertes = list(fus["alertes"]) + analyse.controles()
    fichiers_res = []
    for nom, imp, err in resultats:
        fichiers_res.append({"fichier": nom, "erreur": err,
                             "exercices": [x.exercice for x in (imp.bilan_n, imp.bilan_n1) if x] if imp else [],
                             "alertes": imp.alertes if imp else [], "note_ia": imp.remarques if imp else "",
                             "pages_lues": imp.pages if imp else []})
    return {
        "exercices": fus["exercices"], "fichiers": fichiers_res, "alertes": alertes,
        "entreprise": imports[0].entreprise,
        "bilan_n": asdict(b[0]) if b[0] else None, "bilan_n1": asdict(b[1]) if b[1] else None,
        "bilan_n2": asdict(b[2]) if b[2] else None,
        "cr_n": asdict(c[0]) if c[0] else None, "cr_n1": asdict(c[1]) if c[1] else None,
        "cr_n2": asdict(c[2]) if c[2] else None,
    }


@router.post("/bilan-ouverture")
def bilan_ouverture(corps: Dict[str, Any] = Body(...)):
    """Corps : {projection, capital, stocks_initiaux, creances_initiales, dettes_initiales}."""
    from befa.ouverture import MENTION, proposer
    from befa.projection import ProjectionInput

    proj = None
    if corps.get("projection"):
        try:
            proj = ProjectionInput.from_dict(corps["projection"])
        except (TypeError, KeyError, ValueError) as e:
            raise HTTPException(status_code=422, detail=f"Données financières mal formées : {e}")
    r = proposer(proj, corps.get("capital"), corps.get("stocks_initiaux") or 0, corps.get("creances_initiales") or 0,
                 corps.get("dettes_initiales") or 0)
    return {"bilan": asdict(r.bilan), "hypotheses": r.hypotheses, "alertes": r.alertes, "mention": MENTION,
            "total_actif": r.bilan.total_actif, "total_passif": r.bilan.total_passif}


# ----------------------------------------------------------------------
# Donnees financieres : compte d'exploitation, plan de tresorerie
# ----------------------------------------------------------------------
@router.post("/tableau")
async def importer_tableau(fichier: UploadFile = File(...), nature: str = Form(...), dossier_id: str = Form("")):
    contenu = await _lire(fichier)
    try:
        res = extraire_tableau(fichier.filename or "tableau", contenu, nature)
    except ExtractionError as e:
        _erreur(e)
    except IaError as e:
        _erreur(e, 502)
    stockage.conserver_source(dossier_id or None, fichier.filename or "tableau", contenu, nature, res)
    return res


# ----------------------------------------------------------------------
# Demandes de financement bancaires
# ----------------------------------------------------------------------
@router.post("/courriers/preparer")
def preparer_courriers(corps: Dict[str, Any] = Body(...)):
    """Corps : {dossier, banques: [id...]} -> un courrier modifiable par banque."""
    from befa import courriers

    d, proj = _dossier(corps.get("dossier") or {})
    ids = corps.get("banques") or []
    if not ids:
        raise HTTPException(status_code=422, detail="Sélectionnez au moins une banque.")
    res = []
    for ident in ids:
        try:
            b = referentiels.obtenir("banques", ident)
        except referentiels.ReferentielError as e:
            _erreur(e, 404)
        c = courriers.preparer(d, proj, b)
        c["banque_sigle"] = b.get("sigle", "")
        res.append(c)
    return {"courriers": res}


@router.post("/courriers/generer")
def generer_courriers(corps: Dict[str, Any] = Body(...)):
    """Corps : {dossier, courriers: [...], dossier_id} -> Word + PDF par banque + ZIP."""
    from befa import courriers
    from befa.pdf import PdfError

    d, _ = _dossier(corps.get("dossier") or {})
    sortie = os.path.join(_SORTIES, uuid.uuid4().hex)
    try:
        res = courriers.generer(d, corps.get("courriers") or [], sortie)
    except ValueError as e:
        _erreur(e)
    except PdfError as e:
        _erreur(e, 500)
    tous = [x[k] for x in res["courriers"] for k in ("docx", "pdf") if x[k]] + [res["zip"]]
    urls = _publier(tous)
    version = stockage.archiver_documents(corps.get("dossier_id") or None, "courriers", tous,
                                          {"banques": [x["banque"] for x in res["courriers"]]})
    return {"courriers": [{"banque": x["banque"], "docx": urls[os.path.basename(x["docx"])],
                           "pdf": urls.get(os.path.basename(x["pdf"] or ""))} for x in res["courriers"]],
            "zip": urls[os.path.basename(res["zip"])], "alertes": res["alertes"], "version": version}


# ----------------------------------------------------------------------
# Autres documents
# ----------------------------------------------------------------------
@router.post("/documents/preparer")
def preparer_document(corps: Dict[str, Any] = Body(...)):
    """Corps : {dossier, modele_id, precisions} -> contenu modifiable."""
    from befa import autres_documents

    d, proj = _dossier(corps.get("dossier") or {})
    try:
        modele = referentiels.obtenir("modeles", corps.get("modele_id", ""))
    except referentiels.ReferentielError as e:
        _erreur(e, 404)
    try:
        return autres_documents.preparer(d, proj, modele, corps.get("precisions", ""))
    except IaError as e:
        _erreur(e, 502)


@router.post("/documents/generer")
def generer_document(corps: Dict[str, Any] = Body(...)):
    from befa import autres_documents
    from befa.pdf import PdfError

    d, _ = _dossier(corps.get("dossier") or {})
    contenu = corps.get("contenu") or {}
    if not contenu.get("sections") and not contenu.get("tableaux"):
        raise HTTPException(status_code=422, detail="Document vide : préparez-le d'abord.")
    sortie = os.path.join(_SORTIES, uuid.uuid4().hex)
    try:
        res = autres_documents.generer(d, contenu, sortie)
    except PdfError as e:
        _erreur(e, 500)
    urls = _publier([res["docx"], res["pdf"]])
    version = stockage.archiver_documents(corps.get("dossier_id") or None, contenu.get("modele_id") or "document",
                                          [res["docx"], res["pdf"]], {"titre": contenu.get("titre", "")})
    return {"docx": urls[os.path.basename(res["docx"])], "pdf": urls.get(os.path.basename(res["pdf"] or "")),
            "alertes": res["alertes"], "version": version}


# ----------------------------------------------------------------------
# Dossiers conserves sur le serveur
# ----------------------------------------------------------------------
@router.get("/dossiers")
def lister_dossiers():
    return stockage.lister_dossiers()


@router.post("/dossiers")
def creer_dossier():
    return {"id": stockage.nouvel_id()}


@router.put("/dossiers/{ident}")
def enregistrer_dossier(ident: str, etat: Dict[str, Any] = Body(...)):
    try:
        return stockage.enregistrer_dossier(ident, etat)
    except stockage.StockageError as e:
        _erreur(e)


@router.get("/dossiers/{ident}")
def charger_dossier(ident: str):
    try:
        return stockage.charger_dossier(ident)
    except stockage.StockageError as e:
        _erreur(e, 404)


@router.delete("/dossiers/{ident}")
def supprimer_dossier(ident: str):
    try:
        stockage.supprimer_dossier(ident)
    except stockage.StockageError as e:
        _erreur(e, 404)
    return {"supprime": ident}


@router.get("/dossiers/{ident}/versions")
def versions(ident: str):
    try:
        return {"versions": stockage.lister_versions(ident), "sources": stockage.lister_sources(ident)}
    except stockage.StockageError as e:
        _erreur(e, 404)


@router.get("/dossiers/{ident}/versions/{version}/{nom}")
def telecharger_version(ident: str, version: str, nom: str):
    try:
        chemin = stockage.chemin_document(ident, version, nom)
    except stockage.StockageError as e:
        _erreur(e, 404)
    return FileResponse(chemin, media_type=TYPES.get(nom.rsplit(".", 1)[-1].lower(), "application/octet-stream"),
                        filename=nom)
