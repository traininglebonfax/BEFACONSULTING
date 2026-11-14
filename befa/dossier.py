"""
Dossier client BEFA : tout ce que le consultant saisit ou importe pour
produire un Business Plan.

Deux variantes, calquees sur les modeles du cabinet :
  - "nouveau_projet"       (COREAL FOOD) : presentation, requete, environnement,
                             projet et marche, organisation, bilan d'ouverture,
                             previsions, risques, SWOT, perspectives, impact ;
  - "entreprise_existante" (O'IVOIRE) : presentation, etude financiere N/N-1,
                             besoin et projections, etude de marche.
"""
from dataclasses import asdict, dataclass, field
from typing import List, Optional

from .historique import AnalyseHistorique, BilanCondense, CompteResultatCondense
from .projection import ProjectionInput

VARIANTES = {
    "nouveau_projet": "Nouveau projet (modèle COREAL FOOD)",
    "entreprise_existante": "Entreprise existante - dossier de crédit (modèle O'IVOIRE)",
}


@dataclass
class Identite:
    raison_sociale: str
    sigle: str = ""
    forme_juridique: str = ""  # SA, SARL, SAS, entreprise individuelle...
    capital: Optional[float] = None
    date_creation: str = ""
    rccm: str = ""
    ncc: str = ""
    regime_fiscal: str = ""
    centre_impots: str = ""
    adresse: str = ""
    siege: str = ""
    ville: str = ""
    telephone: str = ""
    email: str = ""
    objet_social: List[str] = field(default_factory=list)
    activite: str = ""  # nature de l'activite, en une ou deux phrases


@dataclass
class Actionnaire:
    nom: str
    montant: float = 0
    nationalite: str = ""
    nature: str = ""  # personne physique / morale

    def part(self, capital: float) -> Optional[float]:
        return self.montant / capital if capital else None


@dataclass
class Dirigeant:
    nom: str
    fonction: str = ""
    profil: str = ""


@dataclass
class ProduitService:
    nom: str
    description: str = ""


@dataclass
class Partenaire:
    """Client ou fournisseur."""
    nom: str
    mode_paiement: str = ""
    delai: str = ""
    localisation: str = ""


@dataclass
class Projet:
    intitule: str
    secteur: str  # ex. "Transformation agro-industrielle - huile de palme"
    localisation: str = ""  # ville, region
    zone_marche: str = "Côte d'Ivoire"
    objet: str = ""
    description: str = ""
    objectifs: List[str] = field(default_factory=list)
    contexte_besoin: str = ""  # entreprise existante : pourquoi ce financement maintenant
    phases: List[str] = field(default_factory=list)  # planning de realisation
    emplois_directs: Optional[int] = None
    emplois_indirects: Optional[int] = None
    concurrents: List[str] = field(default_factory=list)
    mots_cles_marche: List[str] = field(default_factory=list)  # termes etudies dans Google Trends
    notes_consultant: str = ""  # informations libres a exploiter par la redaction


@dataclass
class DossierBEFA:
    variante: str
    entreprise: Identite
    projet: Projet
    actionnaires: List[Actionnaire] = field(default_factory=list)
    dirigeants: List[Dirigeant] = field(default_factory=list)
    gouvernance: str = ""  # organes de gestion (AG, CA, comites...)
    produits: List[ProduitService] = field(default_factory=list)
    clients: List[Partenaire] = field(default_factory=list)
    fournisseurs: List[Partenaire] = field(default_factory=list)
    pool_bancaire: str = ""
    annee_document: str = ""
    projection: Optional[ProjectionInput] = None
    bilan_n: Optional[BilanCondense] = None
    bilan_n1: Optional[BilanCondense] = None
    cr_n: Optional[CompteResultatCondense] = None
    cr_n1: Optional[CompteResultatCondense] = None
    trends_faits: List[str] = field(default_factory=list)
    trends_source: str = ""

    def validate(self) -> List[str]:
        errs = []
        if self.variante not in VARIANTES:
            errs.append(f"Variante inconnue '{self.variante}' (attendu : {', '.join(VARIANTES)}).")
        if not self.entreprise.raison_sociale.strip():
            errs.append("La raison sociale est obligatoire.")
        if not self.projet.intitule.strip():
            errs.append("L'intitulé du projet est obligatoire.")
        if not self.projet.secteur.strip():
            errs.append("Le secteur d'activité est obligatoire.")
        if self.variante == "entreprise_existante" and self.bilan_n is None:
            errs.append("Entreprise existante : importez au moins la liasse du dernier exercice.")
        return errs

    @property
    def analyse_historique(self) -> Optional[AnalyseHistorique]:
        if self.bilan_n is None or self.cr_n is None:
            return None
        return AnalyseHistorique(self.bilan_n, self.bilan_n1, self.cr_n, self.cr_n1)

    def resume_pour_ia(self) -> dict:
        """Donnees qualitatives transmises a l'IA (sans les tableaux chiffres)."""
        d = {
            "variante": VARIANTES.get(self.variante, self.variante),
            "entreprise": asdict(self.entreprise),
            "projet": asdict(self.projet),
            "actionnaires": [asdict(a) for a in self.actionnaires],
            "dirigeants": [asdict(x) for x in self.dirigeants],
            "gouvernance": self.gouvernance,
            "produits": [asdict(p) for p in self.produits],
            "clients": [asdict(c) for c in self.clients],
            "fournisseurs": [asdict(f) for f in self.fournisseurs],
            "pool_bancaire": self.pool_bancaire,
        }
        return d

    @staticmethod
    def from_dict(d: dict) -> "DossierBEFA":
        k = dict(d)
        k["entreprise"] = Identite(**d["entreprise"])
        k["projet"] = Projet(**d["projet"])
        k["actionnaires"] = [Actionnaire(**x) for x in d.get("actionnaires", [])]
        k["dirigeants"] = [Dirigeant(**x) for x in d.get("dirigeants", [])]
        k["produits"] = [ProduitService(**x) for x in d.get("produits", [])]
        k["clients"] = [Partenaire(**x) for x in d.get("clients", [])]
        k["fournisseurs"] = [Partenaire(**x) for x in d.get("fournisseurs", [])]
        if d.get("projection"):
            k["projection"] = ProjectionInput.from_dict(d["projection"])
        for cle in ("bilan_n", "bilan_n1"):
            if d.get(cle):
                k[cle] = BilanCondense.from_dict(d[cle])
        for cle in ("cr_n", "cr_n1"):
            if d.get(cle):
                k[cle] = CompteResultatCondense(**d[cle])
        return DossierBEFA(**k)
