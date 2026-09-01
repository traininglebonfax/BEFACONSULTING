"""
Backend applicatif.

Recoit les reponses du formulaire, appelle le moteur de generation Excel
(engine/), declenche le recalcul LibreOffice, appelle l'IA pour les textes
narratifs (Lot 4, attaches en commentaires de cellule - jamais bloquant si
indisponible), convertit les feuilles en HTML pour l'apercu (Lot 3), et
sert le fichier .xlsx final au telechargement.

Lancer avec:
    uvicorn backend.app:app --reload --port 8000
"""
import os
import tempfile
import uuid

from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field
from typing import List, Optional

from engine.ai_text import AiTextError, generate_narrative_texts
from engine.generator import attach_ai_comments, generate_workbook
from engine.models import ChargeFixe, Investissement, ProjetInput, ProjetInputError, Produit
from engine.preview import workbook_to_preview
from engine.recalc import RecalcError, recalculate_and_check

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
load_dotenv(os.path.join(BASE_DIR, ".env"))

WEB_DIR = os.path.join(BASE_DIR, "web")
TMP_DIR = os.path.join(tempfile.gettempdir(), "business_plan_ia")
os.makedirs(TMP_DIR, exist_ok=True)

# Jobs deja generes + recalcules, prets a etre telecharges sans recalcul
# supplementaire (cf. cahier des charges §5 : apercu puis telechargement du
# meme fichier, jamais deux recalculs distincts).
_JOBS: dict = {}

app = FastAPI(title="Business Plan IA - API")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


# --------------------------------------------------------------------
# Schemas de requete (miroir de engine.models, en Pydantic pour la
# validation HTTP - la validation metier definitive reste dans
# ProjetInput.validate(), jamais dupliquee ici)
# --------------------------------------------------------------------
class ProduitIn(BaseModel):
    nom: str
    prix_unitaire: float
    volume_initial_mensuel: float
    cout_variable_unitaire: Optional[float] = None
    cout_matiere_premiere: Optional[float] = None
    cout_emballage: Optional[float] = None
    cout_production: Optional[float] = None
    unite_volume: str = "unites/mois"


class ChargeFixeIn(BaseModel):
    nom: str
    montant_mensuel: float


class InvestissementIn(BaseModel):
    nom: str
    montant: float
    duree_amortissement_ans: float


class ProjetIn(BaseModel):
    secteur_activite: str
    produits: List[ProduitIn]
    charges_fixes: List[ChargeFixeIn]
    taux_croissance_a2: float
    taux_croissance_a3: float
    investissements: List[InvestissementIn] = Field(default_factory=list)
    saisonnalite: Optional[List[float]] = None

    taux_evolution_charges_a2: float = 0.05
    taux_evolution_charges_a3: float = 0.10
    part_fonds_propres: float = 0.35
    part_emprunt: float = 0.65
    taux_interet_annuel: float = 0.09
    duree_remboursement_mois: int = 48
    taux_is: float = 0.25
    delai_clients_jours: int = 30
    delai_fournisseurs_jours: int = 45
    jours_ouvres_mensuel: int = 24
    capacite_max_mensuelle: Optional[float] = None

    nom_projet: str = ""
    forme_juridique: str = ""
    pays_devise: str = "Cote d'Ivoire / FCFA"
    date_demarrage: str = ""

    description_activite: str = ""
    justification_croissance: str = ""


def _to_projet_input(payload: ProjetIn) -> ProjetInput:
    return ProjetInput(
        secteur_activite=payload.secteur_activite,
        produits=[Produit(**p.model_dump()) for p in payload.produits],
        charges_fixes=[ChargeFixe(**c.model_dump()) for c in payload.charges_fixes],
        taux_croissance_a2=payload.taux_croissance_a2,
        taux_croissance_a3=payload.taux_croissance_a3,
        investissements=[Investissement(**i.model_dump()) for i in payload.investissements],
        saisonnalite=payload.saisonnalite,
        taux_evolution_charges_a2=payload.taux_evolution_charges_a2,
        taux_evolution_charges_a3=payload.taux_evolution_charges_a3,
        part_fonds_propres=payload.part_fonds_propres,
        part_emprunt=payload.part_emprunt,
        taux_interet_annuel=payload.taux_interet_annuel,
        duree_remboursement_mois=payload.duree_remboursement_mois,
        taux_is=payload.taux_is,
        delai_clients_jours=payload.delai_clients_jours,
        delai_fournisseurs_jours=payload.delai_fournisseurs_jours,
        jours_ouvres_mensuel=payload.jours_ouvres_mensuel,
        capacite_max_mensuelle=payload.capacite_max_mensuelle,
        nom_projet=payload.nom_projet,
        forme_juridique=payload.forme_juridique,
        pays_devise=payload.pays_devise,
        date_demarrage=payload.date_demarrage,
        description_activite=payload.description_activite,
        justification_croissance=payload.justification_croissance,
    )


def _generate_and_recalc(payload: ProjetIn) -> tuple:
    """Valide, genere, recalcule le classeur, puis genere les textes narratifs IA
    (Lot 4) et les attache en commentaires de cellule (jamais bloquant : si l'IA
    est indisponible, la generation du Business Plan continue sans les textes).

    Retourne (job_id, out_path, projet, ai_texts, ai_warning).
    """
    projet = _to_projet_input(payload)

    try:
        projet.validate()
    except ProjetInputError as e:
        raise HTTPException(status_code=422, detail=str(e))

    job_id = uuid.uuid4().hex
    out_path = os.path.join(TMP_DIR, f"{job_id}.xlsx")

    try:
        generate_workbook(projet, out_path)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erreur de generation: {e}")

    try:
        recalculate_and_check(out_path)
    except RecalcError as e:
        raise HTTPException(status_code=500, detail=f"Erreur de recalcul: {e}")

    ai_texts = None
    ai_warning = None
    try:
        ai_texts = generate_narrative_texts(projet)
        attach_ai_comments(out_path, ai_texts["description_activite"], ai_texts["justification_croissance"])
    except AiTextError as e:
        ai_warning = str(e)

    return job_id, out_path, projet, ai_texts, ai_warning


@app.post("/api/preview")
def preview(payload: ProjetIn):
    job_id, out_path, projet, ai_texts, ai_warning = _generate_and_recalc(payload)

    filename_base = projet.nom_projet.strip() or projet.secteur_activite.strip() or "business_plan"
    safe_name = "".join(c if c.isalnum() or c in " -_" else "_" for c in filename_base).strip() or "business_plan"
    _JOBS[job_id] = {"path": out_path, "filename": f"{safe_name}.xlsx"}

    sheets = workbook_to_preview(out_path)
    return {"job_id": job_id, "sheets": sheets, "ai_texts": ai_texts, "ai_warning": ai_warning}


@app.post("/api/generate")
def generate(payload: ProjetIn):
    """Conserve pour compatibilite (Lot 2) : genere et telecharge directement,
    sans passer par l'apercu. Inclut aussi les textes IA (Lot 4) en commentaires."""
    job_id, out_path, projet, _ai_texts, _ai_warning = _generate_and_recalc(payload)

    filename_base = projet.nom_projet.strip() or projet.secteur_activite.strip() or "business_plan"
    safe_name = "".join(c if c.isalnum() or c in " -_" else "_" for c in filename_base).strip() or "business_plan"

    return FileResponse(
        out_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=f"{safe_name}.xlsx",
    )


@app.get("/api/download/{job_id}")
def download(job_id: str):
    job = _JOBS.get(job_id)
    if not job or not os.path.exists(job["path"]):
        raise HTTPException(status_code=404, detail="Apercu introuvable ou expire. Regenerez l'apercu.")

    return FileResponse(
        job["path"],
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=job["filename"],
    )


@app.get("/api/health")
def health():
    return {"status": "ok"}


if os.path.isdir(WEB_DIR):
    app.mount("/", StaticFiles(directory=WEB_DIR, html=True), name="web")
