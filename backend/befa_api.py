"""
Routes HTTP du Business Plan BEFA (phase 2).

Lot 2 : POST /api/befa/liasse - import d'une liasse SYSCOHADA (PDF).
Le resultat est un brouillon a VALIDER par le consultant : chiffres
condenses, tableaux d'analyse, alertes de controle, note de lecture de l'IA.
"""
from dataclasses import asdict
from typing import Optional

from fastapi import APIRouter, File, HTTPException, UploadFile

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
