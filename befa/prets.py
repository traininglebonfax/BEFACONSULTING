"""
Echeanciers de prets (CMT, CLT, credit-bail, spot...).

Conventions bancaires retenues (usage UEMOA) :
  - taux periodique proportionnel : taux_annuel * mois_par_periode / 12 ;
  - differe TOTAL : aucune echeance, les interets de la periode sont
    capitalises (ajoutes au capital restant du) ;
  - differe PARTIEL : seuls les interets sont payes ;
  - puis echeances constantes (annuites constantes) sur les periodes
    restantes, ou remboursement in fine (credit spot) ;
  - la duree annoncee INCLUT le differe (ex. COREAL : "84 mois dont 18 mois
    de differe").

Les interets sont une charge de la periode (frais financiers du compte
d'exploitation), qu'ils soient payes ou capitalises.

Positionnement dans le temps : l'annee 1 est la premiere annee
d'exploitation. Un pret debloque en "annee 0" (phase d'investissement) ou en
annee 1 demarre au debut de l'annee 1 ; un pret debloque en annee k >= 1
demarre au debut de l'annee k. `decalage_mois` retarde le demarrage de
quelques mois a l'interieur de cette annee.
"""
from dataclasses import dataclass, field
from typing import Dict, List, Optional

PERIODICITES = {"mensuelle": 1, "trimestrielle": 3, "semestrielle": 6, "annuelle": 12}

TYPES_PRET = {
    "CMT": "Crédit moyen terme",
    "CLT": "Crédit long terme",
    "CREDIT_BAIL": "Crédit-bail",
    "SPOT": "Crédit spot",
    "CCT": "Crédit court terme",
}


class FinanceInputError(ValueError):
    """Donnees financieres invalides (message lisible par le consultant)."""


@dataclass
class Echeance:
    numero: int
    mois: int  # mois de l'echeance, compte depuis le debut de l'annee 1 (1 = janvier de l'annee 1)
    capital_debut: float
    interet: float  # charge d'interet de la periode (payee ou capitalisee)
    interet_paye: float
    amortissement: float  # capital rembourse
    echeance: float  # montant decaisse = interet_paye + amortissement
    capital_fin: float

    @property
    def annee(self) -> int:
        return (self.mois - 1) // 12 + 1


@dataclass
class Pret:
    libelle: str
    montant: float
    taux_annuel: float  # 0.09 pour 9 %
    duree_mois: int  # differe inclus
    type: str = "CMT"
    differe_total_mois: int = 0
    differe_partiel_mois: int = 0
    periodicite: str = "mensuelle"
    annee_deblocage: int = 0  # 0 = phase d'investissement (avant l'annee 1)
    decalage_mois: int = 0
    in_fine: bool = False  # capital rembourse en une fois a la derniere echeance
    objet: str = ""
    garantie: str = ""
    # Un credit spot finance le cycle d'exploitation (tire et rembourse dans
    # l'annee) : comme dans le modele COREAL, il n'apparait pas dans le plan
    # de financement, seuls ses interets pesent sur le compte d'exploitation.
    inclure_plan_financement: Optional[bool] = None

    def __post_init__(self):
        self.type = (self.type or "CMT").upper()
        if self.inclure_plan_financement is None:
            self.inclure_plan_financement = self.type != "SPOT"

    @property
    def libelle_type(self) -> str:
        return TYPES_PRET.get(self.type, self.type)

    @property
    def mois_par_periode(self) -> int:
        return PERIODICITES.get(self.periodicite, 0)

    @property
    def mois_debut(self) -> int:
        """Mois (base 0) du deblocage, depuis le debut de l'annee 1."""
        return (max(self.annee_deblocage, 1) - 1) * 12 + self.decalage_mois

    def validate(self, index: int) -> List[str]:
        nom = f"Prêt #{index} ({self.libelle or 'sans nom'})"
        errs = []
        if not (self.libelle or "").strip():
            errs.append(f"Prêt #{index} : le libellé est obligatoire.")
        if self.type not in TYPES_PRET:
            errs.append(f"{nom} : type inconnu '{self.type}' (attendu : {', '.join(TYPES_PRET)}).")
        if self.montant is None or self.montant <= 0:
            errs.append(f"{nom} : le montant doit être > 0.")
        if self.taux_annuel is None or not (0 <= self.taux_annuel < 1):
            errs.append(f"{nom} : le taux annuel doit être compris entre 0 et 1 (ex. 0.09 pour 9 %).")
        if self.periodicite not in PERIODICITES:
            errs.append(f"{nom} : périodicité inconnue '{self.periodicite}' (attendu : {', '.join(PERIODICITES)}).")
            return errs
        p = self.mois_par_periode
        if not self.duree_mois or self.duree_mois <= 0:
            errs.append(f"{nom} : la durée doit être > 0 mois.")
            return errs
        for label, val in (
            ("la durée", self.duree_mois),
            ("le différé total", self.differe_total_mois),
            ("le différé partiel", self.differe_partiel_mois),
        ):
            if val < 0:
                errs.append(f"{nom} : {label} ne peut pas être négatif.")
            elif val % p:
                errs.append(f"{nom} : {label} ({val} mois) doit être un multiple de la périodicité ({p} mois).")
        if self.differe_total_mois + self.differe_partiel_mois >= self.duree_mois:
            errs.append(f"{nom} : le différé ({self.differe_total_mois + self.differe_partiel_mois} mois) doit être inférieur à la durée ({self.duree_mois} mois).")
        if self.annee_deblocage < 0:
            errs.append(f"{nom} : l'année de déblocage ne peut pas être négative.")
        if not (0 <= self.decalage_mois < 12):
            errs.append(f"{nom} : le décalage de démarrage doit être compris entre 0 et 11 mois.")
        return errs

    def echeancier(self) -> List[Echeance]:
        errs = self.validate(1)
        if errs:
            raise FinanceInputError("\n".join(errs))

        p = self.mois_par_periode
        i = self.taux_annuel * p / 12
        n_total = self.duree_mois // p
        n_dt = self.differe_total_mois // p
        n_dp = self.differe_partiel_mois // p
        n_amort = n_total - n_dt - n_dp

        lignes: List[Echeance] = []
        k = float(self.montant)
        annuite = None
        for num in range(1, n_total + 1):
            mois = self.mois_debut + num * p
            interet = k * i
            if num <= n_dt:
                lignes.append(Echeance(num, mois, k, interet, 0.0, 0.0, 0.0, k + interet))
                k += interet
                continue
            if num <= n_dt + n_dp or (self.in_fine and num < n_total):
                lignes.append(Echeance(num, mois, k, interet, interet, 0.0, interet, k))
                continue
            if self.in_fine:
                amort = k
            else:
                if annuite is None:
                    annuite = k / n_amort if i == 0 else k * i / (1 - (1 + i) ** -n_amort)
                amort = k if num == n_total else annuite - interet
            lignes.append(Echeance(num, mois, k, interet, interet, amort, interet + amort, k - amort))
            k -= amort
        return lignes

    def par_annee(self, nb_annees: int) -> Dict[str, List[float]]:
        """Agrege l'echeancier par annee d'exploitation (1..nb_annees).

        Retourne {"interets", "amortissements", "echeances", "capital_fin"}
        (listes de nb_annees valeurs) et "hors_horizon" : capital restant a
        rembourser apres la derniere annee du plan.
        """
        res = {
            "interets": [0.0] * nb_annees,
            "amortissements": [0.0] * nb_annees,
            "echeances": [0.0] * nb_annees,
            "capital_fin": [0.0] * nb_annees,
        }
        hors_horizon = 0.0
        ech = self.echeancier()
        for e in ech:
            a = e.annee
            if a > nb_annees:
                hors_horizon += e.amortissement
                continue
            res["interets"][a - 1] += e.interet
            res["amortissements"][a - 1] += e.amortissement
            res["echeances"][a - 1] += e.echeance
        # capital restant du en fin de chaque annee
        debut = max(self.annee_deblocage, 1)
        for a in range(1, nb_annees + 1):
            if a < debut:
                res["capital_fin"][a - 1] = 0.0 if self.annee_deblocage >= 1 else float(self.montant)
                continue
            dans_annee = [e for e in ech if e.annee <= a]
            res["capital_fin"][a - 1] = dans_annee[-1].capital_fin if dans_annee else float(self.montant)
        res["hors_horizon"] = hors_horizon
        return res

    @staticmethod
    def from_dict(d: dict) -> "Pret":
        return Pret(**d)
