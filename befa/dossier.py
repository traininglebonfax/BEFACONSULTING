"""
Dossier client BEFA : tout ce que le consultant saisit ou importe pour
produire un Business Plan.

Deux profils de demandeur (cles internes conservees pour la compatibilite
des dossiers deja enregistres) :
  - "nouveau_projet"       = STARTUP : entreprise en creation, sans historique
                             comptable ; aucun champ "Entreprise" obligatoire,
                             bilan d'ouverture propose automatiquement ;
  - "entreprise_existante" = PME, ETI, GE : RCCM obligatoire, trois derniers
                             bilans, compte d'exploitation, plan de tresorerie.
Les deux suivent le plan du modele de reference "Dossier de credit O'IVOIRE".
"""
from dataclasses import asdict, dataclass, field, fields
from typing import Dict, List, Optional

from .historique import AnalyseHistorique, BilanCondense, CompteResultatCondense
from .projection import ProjectionInput

VARIANTES = {
    "nouveau_projet": "StartUP (entreprise en création, sans historique comptable)",
    "entreprise_existante": "PME, ETI, GE (entreprise existante sollicitant un financement)",
}
ALIAS_VARIANTES = {"startup": "nouveau_projet", "pme_eti_ge": "entreprise_existante", "pme": "entreprise_existante"}


def _construire(cls, d: Optional[dict]):
    """Instancie une dataclass en ignorant les cles inconnues (dossiers produits par une autre version)."""
    noms = {f.name for f in fields(cls)}
    return cls(**{k: v for k, v in (d or {}).items() if k in noms})


@dataclass
class Identite:
    raison_sociale: str = ""
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
    domaine_activite: str = ""
    historique: str = ""  # historique de la creation (dates cles, evolutions)


@dataclass
class Actionnaire:
    nom: str
    montant: float = 0
    nationalite: str = ""
    nature: str = ""  # personne physique / morale
    nombre_parts: Optional[float] = None  # parts sociales ou actions
    pourcentage: Optional[float] = None  # part du capital (0-1) lue dans les statuts
    date_naissance: str = ""
    qualite: str = ""  # associe gerant, actionnaire fondateur...

    def part(self, capital: float) -> Optional[float]:
        if self.pourcentage is not None:
            return self.pourcentage
        return (self.montant or 0) / capital if capital else None


@dataclass
class Dirigeant:
    nom: str
    fonction: str = ""
    profil: str = ""
    qualite: str = ""  # qualite juridique : gerant statutaire, PDG, administrateur...
    nationalite: str = ""
    date_naissance: str = ""
    adresse: str = ""


@dataclass
class ProduitService:
    nom: str
    description: str = ""


@dataclass
class Partenaire:
    """Client ou fournisseur (saisi ou importe d'un fichier Excel)."""
    nom: str
    mode_paiement: str = ""
    delai: str = ""
    localisation: str = ""
    secteur: str = ""  # secteur / domaine d'activite du partenaire
    nature: str = ""  # fournisseur : nature des biens ou services fournis
    montant: Optional[float] = None  # client : CA realise ; fournisseur : achats annuels (FCFA)
    part: Optional[float] = None  # part dans le CA global / les achats globaux (0-1)


@dataclass
class LigneImportee:
    libelle: str
    categorie: str = "autre"  # cf. befa.extraction.CATEGORIES
    valeurs: List[Optional[float]] = field(default_factory=list)
    total: Optional[float] = None
    style: str = ""  # "", "sous_total", "total"


@dataclass
class TableauImporte:
    """Compte d'exploitation ou plan de tresorerie importe (lu par l'IA, valide par le consultant)."""
    titre: str
    periodes: List[str] = field(default_factory=list)
    lignes: List[LigneImportee] = field(default_factory=list)
    fichier: str = ""
    nature: str = "previsionnel"  # "previsionnel" ou "historique"

    @staticmethod
    def from_dict(d: dict) -> "TableauImporte":
        k = {x: v for x, v in d.items() if x in {f.name for f in fields(TableauImporte)}}
        k["lignes"] = [_construire(LigneImportee, l) for l in d.get("lignes", [])]
        return TableauImporte(**k)


@dataclass
class IdentiteVisuelle:
    """Logo et papier a en-tete du client, en data URL (data:image/png;base64,...)."""
    logo: str = ""
    entete: str = ""  # image, document Word (.docx) ou PDF
    entete_nom: str = ""


@dataclass
class Projet:
    intitule: str
    secteur: str  # libelle du referentiel (befa.secteurs), eventuellement precise
    secteur_code: str = ""  # code du referentiel befa.secteurs
    sous_secteur: str = ""
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
    # PME / ETI / GE : 3e exercice (2 liasses importees couvrent N, N-1 et N-2)
    bilan_n2: Optional[BilanCondense] = None
    cr_n2: Optional[CompteResultatCondense] = None
    # STARTUP : bilan d'ouverture PREVISIONNEL (propose par le programme, modifie et valide par le consultant)
    bilan_ouverture: Optional[BilanCondense] = None
    bilan_ouverture_valide: bool = False
    # Documents importes (page "Donnees financieres")
    compte_exploitation_importe: Optional[TableauImporte] = None
    plan_tresorerie: Optional[TableauImporte] = None
    identite_visuelle: IdentiteVisuelle = field(default_factory=IdentiteVisuelle)

    @property
    def startup(self) -> bool:
        return self.variante == "nouveau_projet"

    @property
    def nom(self) -> str:
        """Nom a afficher : sigle, raison sociale, ou a defaut l'intitule du projet (startup)."""
        e = self.entreprise
        return (e.sigle or e.raison_sociale or "").strip() or self.projet.intitule.strip() or "Le promoteur"

    @property
    def raison_sociale(self) -> str:
        return self.entreprise.raison_sociale.strip() or self.nom

    def validate(self) -> List[str]:
        errs = []
        if self.variante not in VARIANTES:
            errs.append(f"Type de dossier inconnu '{self.variante}' (attendu : StartUP ou PME, ETI, GE).")
        # STARTUP : aucun champ de la page Entreprise n'est obligatoire.
        if not self.startup and not self.entreprise.raison_sociale.strip():
            errs.append("PME, ETI, GE : la raison sociale est obligatoire (importez le RCCM pour la remplir).")
        if not self.projet.intitule.strip():
            errs.append("L'intitulé du projet est obligatoire.")
        if not self.projet.secteur.strip():
            errs.append("Le secteur d'activité est obligatoire.")
        if self.variante == "entreprise_existante" and self.bilan_n is None:
            errs.append("PME, ETI, GE : importez au moins le bilan du dernier exercice.")
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
        if self.bilan_ouverture is not None:
            b = self.bilan_ouverture
            d["bilan_ouverture_previsionnel"] = {
                "nature": "ESTIMATION prévisionnelle, pas une donnée comptable historique",
                "total": b.total_actif, "capital": b.capital, "apports_compte_courant": b.compte_courant_associes,
                "emprunts": b.dettes_financieres, "disponibilites": b.tresorerie_actif,
            }
        return d

    @staticmethod
    def from_dict(d: dict) -> "DossierBEFA":
        noms = {f.name for f in fields(DossierBEFA)}
        k = {x: v for x, v in d.items() if x in noms}
        k["variante"] = ALIAS_VARIANTES.get(d.get("variante", ""), d.get("variante", ""))
        k["entreprise"] = _construire(Identite, d.get("entreprise"))
        k["projet"] = _construire(Projet, d["projet"])
        for cle, cls in (("actionnaires", Actionnaire), ("dirigeants", Dirigeant), ("produits", ProduitService),
                         ("clients", Partenaire), ("fournisseurs", Partenaire)):
            k[cle] = [_construire(cls, x) for x in d.get(cle) or [] if str((x or {}).get("nom") or "").strip()]
        k["projection"] = ProjectionInput.from_dict(d["projection"]) if d.get("projection") else None
        for cle in ("bilan_n", "bilan_n1", "bilan_n2", "bilan_ouverture"):
            k[cle] = BilanCondense.from_dict(d[cle]) if d.get(cle) else None
        for cle in ("cr_n", "cr_n1", "cr_n2"):
            k[cle] = _construire(CompteResultatCondense, d[cle]) if d.get(cle) else None
        for cle in ("compte_exploitation_importe", "plan_tresorerie"):
            k[cle] = TableauImporte.from_dict(d[cle]) if d.get(cle) else None
        k["identite_visuelle"] = _construire(IdentiteVisuelle, d.get("identite_visuelle"))
        return DossierBEFA(**k)
