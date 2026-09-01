"""
Modele de donnees d'entree du moteur de generation de Business Plan.

Ce module ne contient aucune logique financiere : il decrit uniquement
la forme des donnees qu'un projet doit fournir pour que le moteur puisse
les ecrire dans les cellules de saisie du gabarit Excel (model/Business_Plan_Template.xlsx).

Capacites du gabarit (voir scripts/restructure_template.py) :
  - jusqu'a 12 produits/services
  - jusqu'a 20 charges fixes
  - jusqu'a 15 investissements
"""
from dataclasses import dataclass, field
from typing import List, Optional

MAX_PRODUITS = 12
MAX_CHARGES = 20
MAX_INVESTISSEMENTS = 15


class ProjetInputError(ValueError):
    """Erreur de validation des donnees d'entree d'un projet."""


@dataclass
class Produit:
    nom: str
    prix_unitaire: float
    volume_initial_mensuel: float
    # cout_variable_unitaire peut etre fourni directement (formulaire simple),
    # ou omis si les 3 sous-postes ci-dessous sont fournis (formulaire "Detailler") :
    # le moteur additionne alors automatiquement matiere premiere + emballage + production.
    cout_variable_unitaire: Optional[float] = None
    cout_matiere_premiere: Optional[float] = None
    cout_emballage: Optional[float] = None
    cout_production: Optional[float] = None
    unite_volume: str = "unites/mois"

    def __post_init__(self):
        if self.cout_variable_unitaire is None:
            sous_postes = [self.cout_matiere_premiere, self.cout_emballage, self.cout_production]
            if any(v is not None for v in sous_postes):
                self.cout_variable_unitaire = sum(v or 0 for v in sous_postes)

    def validate(self, index: int) -> List[str]:
        errs = []
        if not self.nom or not self.nom.strip():
            errs.append(f"Produit #{index}: le nom est obligatoire.")
        if self.prix_unitaire is None or self.prix_unitaire <= 0:
            errs.append(f"Produit #{index} ({self.nom}): le prix unitaire doit etre > 0.")
        if self.volume_initial_mensuel is None or self.volume_initial_mensuel < 0:
            errs.append(f"Produit #{index} ({self.nom}): le volume initial mensuel doit etre >= 0.")
        if self.cout_variable_unitaire is None or self.cout_variable_unitaire < 0:
            errs.append(
                f"Produit #{index} ({self.nom}): le cout variable unitaire doit etre fourni "
                "(directement, ou via cout_matiere_premiere/cout_emballage/cout_production) et etre >= 0."
            )
        return errs


@dataclass
class ChargeFixe:
    nom: str
    montant_mensuel: float

    def validate(self, index: int) -> List[str]:
        errs = []
        if not self.nom or not self.nom.strip():
            errs.append(f"Charge fixe #{index}: le nom est obligatoire.")
        if self.montant_mensuel is None or self.montant_mensuel < 0:
            errs.append(f"Charge fixe #{index} ({self.nom}): le montant mensuel doit etre >= 0.")
        return errs


@dataclass
class Investissement:
    nom: str
    montant: float
    duree_amortissement_ans: float

    def validate(self, index: int) -> List[str]:
        errs = []
        if not self.nom or not self.nom.strip():
            errs.append(f"Investissement #{index}: le nom est obligatoire.")
        if self.montant is None or self.montant < 0:
            errs.append(f"Investissement #{index} ({self.nom}): le montant doit etre >= 0.")
        if self.duree_amortissement_ans is None or self.duree_amortissement_ans <= 0:
            errs.append(f"Investissement #{index} ({self.nom}): la duree d'amortissement doit etre > 0.")
        return errs


@dataclass
class ProjetInput:
    # --- Champs "a completer" (specifiques au projet) ---
    secteur_activite: str
    produits: List[Produit]
    charges_fixes: List[ChargeFixe]
    taux_croissance_a2: float
    taux_croissance_a3: float
    investissements: List[Investissement] = field(default_factory=list)
    saisonnalite: Optional[List[float]] = None  # 12 coefficients (mois 1 a 12), defaut = [1.0]*12

    # --- Champs "mode avance" (pre-remplis, modifiables) ---
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
    capacite_max_mensuelle: Optional[float] = None  # None = pas de plafond (grande valeur par defaut)

    # --- Metadonnees projet : AUCUNE cellule correspondante dans le classeur
    # (verifie sur l'original et sur le gabarit restructure). Utilisees pour le
    # nom de fichier de sortie et comme contexte du futur texte IA (Lot 4) -
    # jamais ecrites dans une feuille.
    nom_projet: str = ""
    forme_juridique: str = ""
    pays_devise: str = "Cote d'Ivoire / FCFA"
    date_demarrage: str = ""

    # --- Champs "generes par l'IA" (stubs en attendant le Lot 4) ---
    description_activite: str = ""
    justification_croissance: str = ""

    def validate(self) -> None:
        errs: List[str] = []

        if not self.secteur_activite or not self.secteur_activite.strip():
            errs.append("Le secteur d'activite est obligatoire.")

        if not self.produits:
            errs.append("Au moins un produit/service est obligatoire.")
        elif len(self.produits) > MAX_PRODUITS:
            errs.append(f"Trop de produits ({len(self.produits)}) : le gabarit supporte au maximum {MAX_PRODUITS}.")
        for i, p in enumerate(self.produits, start=1):
            errs.extend(p.validate(i))

        if not self.charges_fixes:
            errs.append("Au moins une charge fixe est obligatoire.")
        elif len(self.charges_fixes) > MAX_CHARGES:
            errs.append(f"Trop de charges fixes ({len(self.charges_fixes)}) : le gabarit supporte au maximum {MAX_CHARGES}.")
        for i, c in enumerate(self.charges_fixes, start=1):
            errs.extend(c.validate(i))

        if len(self.investissements) > MAX_INVESTISSEMENTS:
            errs.append(f"Trop d'investissements ({len(self.investissements)}) : le gabarit supporte au maximum {MAX_INVESTISSEMENTS}.")
        for i, inv in enumerate(self.investissements, start=1):
            errs.extend(inv.validate(i))

        if self.saisonnalite is not None and len(self.saisonnalite) != 12:
            errs.append("La saisonnalite doit comporter exactement 12 coefficients (un par mois).")

        if abs(self.part_fonds_propres + self.part_emprunt - 1.0) > 1e-6:
            errs.append("La somme (part fonds propres + part emprunt) doit valoir 1.0 (100%).")

        if errs:
            raise ProjetInputError("Donnees de projet invalides:\n- " + "\n- ".join(errs))

    @staticmethod
    def from_dict(d: dict) -> "ProjetInput":
        produits = [Produit(**p) for p in d.get("produits", [])]
        charges = [ChargeFixe(**c) for c in d.get("charges_fixes", [])]
        invests = [Investissement(**i) for i in d.get("investissements", [])]
        kwargs = dict(d)
        kwargs["produits"] = produits
        kwargs["charges_fixes"] = charges
        kwargs["investissements"] = invests
        return ProjetInput(**kwargs)
