"""
Etats financiers previsionnels au format BEFA :

  - compte d'exploitation previsionnel (CA -> Marge brute -> Valeur ajoutee
    -> EBE -> Resultat d'exploitation -> Resultat net -> CAF) ;
  - plan de financement (annee 0 d'investissement + annees d'exploitation) ;
  - ratio de couverture de la dette tel que defini par BEFA :
    free cash-flow (RN + dotations) / frais financiers,
    complete par le taux de couverture du service de la dette
    (CAF / (interets + capital rembourse)) attendu par les banques.

Tous les montants sont arrondis au franc CFA le plus proche, ligne par ligne.
"""
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence, Union

from .prets import FinanceInputError, Pret

Nombre = Union[int, float]

RUBRIQUES_CHARGES = {
    "achats_matieres": "Achats de matières premières",
    "autres_achats": "Autres achats",
    "transport": "Transport",
    "services_exterieurs": "Services extérieurs",
    "impots_taxes": "Impôts et taxes",
    "salaires": "Salaires",
}


def _serie(valeur, n: int, nom: str) -> List[float]:
    """Accepte un nombre (constant sur la periode) ou une liste de n valeurs."""
    if isinstance(valeur, (int, float)):
        return [float(valeur)] * n
    valeurs = list(valeur)
    if len(valeurs) != n:
        raise FinanceInputError(f"{nom} : {len(valeurs)} valeur(s) fournie(s), {n} attendue(s) (une par année).")
    return [float(v or 0) for v in valeurs]


def _r(x: float) -> int:
    return int(round(x))


@dataclass
class LigneProduit:
    """Produit ou service vendu : quantites x prix, ou montants annuels directs."""
    libelle: str
    quantites: Optional[Sequence[Nombre]] = None
    prix_unitaire: Optional[Union[Nombre, Sequence[Nombre]]] = None
    montants: Optional[Sequence[Nombre]] = None
    unite: str = ""

    def calculer(self, n: int) -> List[int]:
        if self.montants is not None:
            return [_r(v) for v in _serie(self.montants, n, f"Produit '{self.libelle}'")]
        if self.quantites is None or self.prix_unitaire is None:
            raise FinanceInputError(f"Produit '{self.libelle}' : fournir soit les montants annuels, soit les quantités et le prix unitaire.")
        q = _serie(self.quantites, n, f"Produit '{self.libelle}' (quantités)")
        p = _serie(self.prix_unitaire, n, f"Produit '{self.libelle}' (prix)")
        return [_r(a * b) for a, b in zip(q, p)]


@dataclass
class LigneCharge:
    """Charge d'exploitation rattachee a une rubrique du compte d'exploitation.

    Trois modes, exclusifs :
      - montants : un montant par annee (ou un montant constant) ;
      - pct_ca   : pourcentage du chiffre d'affaires (constant ou par annee) ;
      - quantites x cout_unitaire (ex. tonnes de regimes x cout a la tonne).
    """
    libelle: str
    rubrique: str
    montants: Optional[Union[Nombre, Sequence[Nombre]]] = None
    pct_ca: Optional[Union[Nombre, Sequence[Nombre]]] = None
    quantites: Optional[Sequence[Nombre]] = None
    cout_unitaire: Optional[Union[Nombre, Sequence[Nombre]]] = None

    def calculer(self, n: int, ca: List[int]) -> List[int]:
        nom = f"Charge '{self.libelle}'"
        if self.rubrique not in RUBRIQUES_CHARGES:
            raise FinanceInputError(f"{nom} : rubrique inconnue '{self.rubrique}' (attendu : {', '.join(RUBRIQUES_CHARGES)}).")
        modes = [self.montants is not None, self.pct_ca is not None, self.quantites is not None]
        if sum(modes) != 1:
            raise FinanceInputError(f"{nom} : renseigner un seul mode de calcul (montants, % du CA, ou quantités x coût unitaire).")
        if self.montants is not None:
            return [_r(v) for v in _serie(self.montants, n, nom)]
        if self.pct_ca is not None:
            taux = _serie(self.pct_ca, n, nom)
            return [_r(c * t) for c, t in zip(ca, taux)]
        if self.cout_unitaire is None:
            raise FinanceInputError(f"{nom} : le coût unitaire est obligatoire avec les quantités.")
        q = _serie(self.quantites, n, f"{nom} (quantités)")
        cu = _serie(self.cout_unitaire, n, f"{nom} (coût unitaire)")
        return [_r(a * b) for a, b in zip(q, cu)]


@dataclass
class Investissement:
    libelle: str
    montant: float
    annee: int = 0  # 0 = phase d'investissement, k = acquis au debut de l'annee k
    duree_amortissement_ans: Optional[float] = None  # None ou 0 = non amortissable (terrain...)
    rubrique: str = ""  # regroupement libre pour le recapitulatif (ex. "Génie civil")
    details: List[str] = field(default_factory=list)  # sous-postes affiches sous la rubrique
    financement: str = ""  # "apport", "emprunt", "credit-bail"... (affichage)

    def dotations(self, n: int) -> List[float]:
        res = [0.0] * n
        if not self.duree_amortissement_ans:
            return res
        annuite = self.montant / self.duree_amortissement_ans
        debut = max(self.annee, 1)
        restant = float(self.montant)
        for a in range(debut, n + 1):
            d = min(annuite, restant)
            res[a - 1] = d
            restant -= d
            if restant <= 1e-6:
                break
        return res


@dataclass
class Apport:
    libelle: str
    montant: float
    annee: int = 0
    nature: str = "apport"  # apport, subvention, compte courant...


@dataclass
class ProjectionInput:
    annees: List[str]  # libelles des colonnes : ["Année 1", ...] ou ["2026", ...]
    produits: List[LigneProduit]
    charges: List[LigneCharge] = field(default_factory=list)
    investissements: List[Investissement] = field(default_factory=list)
    prets: List[Pret] = field(default_factory=list)
    apports: List[Apport] = field(default_factory=list)
    taux_is: float = 0.25  # impot sur les societes, Cote d'Ivoire

    # Entreprise existante / ajustements du consultant (une valeur par annee)
    dotations_existantes: Optional[Sequence[Nombre]] = None  # amortissements des actifs deja en place
    dotations_force: Optional[Sequence[Nombre]] = None  # remplace le calcul des dotations
    autres_frais_financiers: Optional[Sequence[Nombre]] = None  # interets de dettes deja en cours
    remboursements_existants: Optional[Sequence[Nombre]] = None  # capital de dettes deja en cours
    impots_sur_benefice_force: Optional[Sequence[Nombre]] = None  # ex. impot minimum forfaitaire

    # Besoin en fonds de roulement : en jours, ou augmentation saisie directement
    jours_clients: float = 0
    jours_stocks: float = 0  # en jours d'achats de matieres
    jours_fournisseurs: float = 0  # en jours d'achats (matieres + autres achats)
    bfr_initial: float = 0  # BFR existant en debut d'annee 1 (entreprise existante)
    augmentation_bfr_force: Optional[Sequence[Nombre]] = None

    tresorerie_initiale: float = 0

    @property
    def n(self) -> int:
        return len(self.annees)

    def validate(self) -> None:
        errs = []
        if not (1 <= self.n <= 15):
            errs.append("La projection doit couvrir entre 1 et 15 années.")
        if not self.produits:
            errs.append("Au moins un produit ou service est obligatoire.")
        if not (0 <= self.taux_is < 1):
            errs.append("Le taux d'impôt sur les sociétés doit être compris entre 0 et 1.")
        for i, p in enumerate(self.prets, start=1):
            errs.extend(p.validate(i))
        for i, inv in enumerate(self.investissements, start=1):
            if inv.montant is None or inv.montant < 0:
                errs.append(f"Investissement #{i} ({inv.libelle}) : le montant doit être >= 0.")
            if not (0 <= inv.annee <= self.n):
                errs.append(f"Investissement #{i} ({inv.libelle}) : année {inv.annee} hors de la période (0 à {self.n}).")
        for a in self.apports:
            if not (0 <= a.annee <= self.n):
                errs.append(f"Apport '{a.libelle}' : année {a.annee} hors de la période (0 à {self.n}).")
        for p in self.prets:
            if p.annee_deblocage > self.n:
                errs.append(f"Prêt '{p.libelle}' : année de déblocage {p.annee_deblocage} hors de la période.")
        if errs:
            raise FinanceInputError("Données financières invalides :\n- " + "\n- ".join(errs))

    @staticmethod
    def from_dict(d: dict) -> "ProjectionInput":
        k = dict(d)
        k["produits"] = [LigneProduit(**x) for x in d.get("produits", [])]
        k["charges"] = [LigneCharge(**x) for x in d.get("charges", [])]
        k["investissements"] = [Investissement(**x) for x in d.get("investissements", [])]
        k["prets"] = [Pret(**x) for x in d.get("prets", [])]
        k["apports"] = [Apport(**x) for x in d.get("apports", [])]
        return ProjectionInput(**k)


# ----------------------------------------------------------------------
# Resultats
# ----------------------------------------------------------------------
@dataclass
class Ligne:
    """Ligne de tableau prete a etre rendue (PDF, HTML)."""
    libelle: str
    valeurs: List[Optional[float]]
    style: str = ""  # "" | "total" | "sous_total" | "detail" | "ratio" | "pct"


@dataclass
class CompteExploitation:
    annees: List[str]
    ca_par_produit: Dict[str, List[int]]
    chiffre_affaires: List[int]
    charges_par_rubrique: Dict[str, List[int]]
    charges_detail: Dict[str, Dict[str, List[int]]]
    marge_brute: List[int]
    valeur_ajoutee: List[int]
    ebe: List[int]
    dotations: List[int]
    resultat_exploitation: List[int]
    frais_financiers: List[int]
    resultat_avant_impot: List[int]
    impot_sur_benefice: List[int]
    resultat_net: List[int]
    caf: List[int]

    def lignes(self) -> List[Ligne]:
        c = self.charges_par_rubrique
        out = [Ligne("Chiffre d'affaires total", self.chiffre_affaires, "total")]
        if len(self.ca_par_produit) > 1:
            out[0:0] = [Ligne(nom, v, "detail") for nom, v in self.ca_par_produit.items()]
        if any(c["achats_matieres"]):
            out.append(Ligne(RUBRIQUES_CHARGES["achats_matieres"], c["achats_matieres"]))
        out.append(Ligne("Marge brute", self.marge_brute, "sous_total"))
        for r in ("autres_achats", "transport", "services_exterieurs", "impots_taxes"):
            out.append(Ligne(RUBRIQUES_CHARGES[r], c[r]))
        out += [
            Ligne("Valeur ajoutée", self.valeur_ajoutee, "sous_total"),
            Ligne("Salaires", c["salaires"]),
            Ligne("Excédent Brut d'Exploitation", self.ebe, "sous_total"),
            Ligne("Dotations aux amortissements", self.dotations),
            Ligne("Résultat d'exploitation", self.resultat_exploitation, "sous_total"),
            Ligne("Frais financiers", self.frais_financiers),
            Ligne("Impôt sur les bénéfices", self.impot_sur_benefice),
            Ligne("Résultat net", self.resultat_net, "total"),
            Ligne("CAF", self.caf, "total"),
        ]
        return out


@dataclass
class PlanFinancement:
    colonnes: List[str]  # ["Année 0", "Année 1", ...] (Année 0 seulement si utile)
    lignes_ressources: List[Ligne]
    flux_entrants: List[int]
    lignes_emplois: List[Ligne]
    flux_sortants: List[int]
    flux_nets: List[int]
    cumul: List[int]

    def lignes(self) -> List[Ligne]:
        return (
            self.lignes_ressources
            + [Ligne("Flux entrants", self.flux_entrants, "sous_total")]
            + self.lignes_emplois
            + [
                Ligne("Flux sortants", self.flux_sortants, "sous_total"),
                Ligne("Flux nets", self.flux_nets, "total"),
                Ligne("Cumul flux nets", self.cumul, "total"),
            ]
        )


@dataclass
class CouvertureDette:
    annees: List[str]
    resultat_net: List[int]
    amortissements: List[int]
    free_cash_flow: List[int]
    frais_financiers: List[int]
    ratio_couverture: List[Optional[float]]  # definition BEFA : FCF / frais financiers
    service_dette: List[int]  # interets + capital rembourse
    dscr: List[Optional[float]]  # CAF / service de la dette

    def lignes(self) -> List[Ligne]:
        return [
            Ligne("Résultat net", self.resultat_net),
            Ligne("Amortissements", self.amortissements),
            Ligne("Free cash-flow", self.free_cash_flow, "sous_total"),
            Ligne("Frais financiers", self.frais_financiers),
            Ligne("Ratio de couverture de la dette", self.ratio_couverture, "ratio"),
        ]

    def lignes_service_dette(self) -> List[Ligne]:
        return [
            Ligne("CAF", self.free_cash_flow),
            Ligne("Service de la dette (intérêts + capital)", self.service_dette),
            Ligne("Taux de couverture du service de la dette", self.dscr, "ratio"),
        ]


@dataclass
class Projection:
    entree: ProjectionInput
    compte: CompteExploitation
    plan: PlanFinancement
    couverture: CouvertureDette
    bfr: List[int]
    augmentation_bfr: List[int]
    echeanciers: Dict[str, dict]  # libelle du pret -> agregats annuels
    indicateurs: Dict[str, List[Optional[float]]]

    @property
    def cout_total_investissements(self) -> int:
        return _r(sum(i.montant for i in self.entree.investissements))

    @property
    def total_prets(self) -> int:
        return _r(sum(p.montant for p in self.entree.prets))

    @property
    def total_prets_investissement(self) -> int:
        return _r(sum(p.montant for p in self.entree.prets if p.inclure_plan_financement))

    @property
    def total_apports(self) -> int:
        return _r(sum(a.montant for a in self.entree.apports))


# ----------------------------------------------------------------------
# Calcul
# ----------------------------------------------------------------------
def _ratio(a: float, b: float) -> Optional[float]:
    return None if not b else a / b


def calculer_projection(entree: ProjectionInput) -> Projection:
    entree.validate()
    n = entree.n

    # --- Chiffre d'affaires
    ca_par_produit = {}
    for p in entree.produits:
        nom = p.libelle
        while nom in ca_par_produit:
            nom += " "
        ca_par_produit[nom] = p.calculer(n)
    ca = [sum(v[a] for v in ca_par_produit.values()) for a in range(n)]

    # --- Charges par rubrique
    detail: Dict[str, Dict[str, List[int]]] = {r: {} for r in RUBRIQUES_CHARGES}
    for ch in entree.charges:
        valeurs = ch.calculer(n, ca)  # leve FinanceInputError si la rubrique est inconnue
        existant = detail[ch.rubrique].get(ch.libelle)
        detail[ch.rubrique][ch.libelle] = valeurs if existant is None else [a + b for a, b in zip(existant, valeurs)]
    par_rubrique = {r: [sum(v[a] for v in d.values()) for a in range(n)] for r, d in detail.items()}

    marge_brute = [ca[a] - par_rubrique["achats_matieres"][a] for a in range(n)]
    va = [
        marge_brute[a]
        - par_rubrique["autres_achats"][a]
        - par_rubrique["transport"][a]
        - par_rubrique["services_exterieurs"][a]
        - par_rubrique["impots_taxes"][a]
        for a in range(n)
    ]
    ebe = [va[a] - par_rubrique["salaires"][a] for a in range(n)]

    # --- Dotations
    if entree.dotations_force is not None:
        dotations = [_r(v) for v in _serie(entree.dotations_force, n, "Dotations aux amortissements")]
    else:
        dot = [0.0] * n
        for inv in entree.investissements:
            for a, d in enumerate(inv.dotations(n)):
                dot[a] += d
        if entree.dotations_existantes is not None:
            for a, d in enumerate(_serie(entree.dotations_existantes, n, "Dotations existantes")):
                dot[a] += d
        dotations = [_r(v) for v in dot]
    re_ = [ebe[a] - dotations[a] for a in range(n)]

    # --- Prets
    echeanciers = {}
    interets = [0.0] * n
    # Capital rembourse : hors credits de campagne (spot), tires et rembourses
    # dans le cycle d'exploitation ; seuls leurs interets pesent sur le resultat
    # et sur le service de la dette.
    capital_rembourse_plan = [0.0] * n
    for p in entree.prets:
        agg = p.par_annee(n)
        nom = p.libelle
        while nom in echeanciers:
            nom += " "
        echeanciers[nom] = agg
        for a in range(n):
            interets[a] += agg["interets"][a]
            if p.inclure_plan_financement:
                capital_rembourse_plan[a] += agg["amortissements"][a]
    if entree.autres_frais_financiers is not None:
        for a, v in enumerate(_serie(entree.autres_frais_financiers, n, "Autres frais financiers")):
            interets[a] += v
    if entree.remboursements_existants is not None:
        for a, v in enumerate(_serie(entree.remboursements_existants, n, "Remboursements existants")):
            capital_rembourse_plan[a] += v
    frais_fin = [_r(v) for v in interets]

    rai = [re_[a] - frais_fin[a] for a in range(n)]
    if entree.impots_sur_benefice_force is not None:
        impot = [_r(v) for v in _serie(entree.impots_sur_benefice_force, n, "Impôt sur les bénéfices")]
    else:
        impot = [_r(max(0, r) * entree.taux_is) for r in rai]
    rn = [rai[a] - impot[a] for a in range(n)]
    caf = [rn[a] + dotations[a] for a in range(n)]

    compte = CompteExploitation(
        annees=list(entree.annees),
        ca_par_produit=ca_par_produit,
        chiffre_affaires=ca,
        charges_par_rubrique=par_rubrique,
        charges_detail=detail,
        marge_brute=marge_brute,
        valeur_ajoutee=va,
        ebe=ebe,
        dotations=dotations,
        resultat_exploitation=re_,
        frais_financiers=frais_fin,
        resultat_avant_impot=rai,
        impot_sur_benefice=impot,
        resultat_net=rn,
        caf=caf,
    )

    # --- BFR
    if entree.augmentation_bfr_force is not None:
        aug_bfr = [_r(v) for v in _serie(entree.augmentation_bfr_force, n, "Augmentation du BFR")]
        bfr, niveau = [], entree.bfr_initial
        for v in aug_bfr:
            niveau += v
            bfr.append(_r(niveau))
    else:
        bfr = []
        for a in range(n):
            achats = par_rubrique["achats_matieres"][a]
            achats_fournisseurs = achats + par_rubrique["autres_achats"][a]
            bfr.append(_r(
                ca[a] * entree.jours_clients / 360
                + achats * entree.jours_stocks / 360
                - achats_fournisseurs * entree.jours_fournisseurs / 360
            ))
        precedent = [_r(entree.bfr_initial)] + bfr[:-1]
        aug_bfr = [bfr[a] - precedent[a] for a in range(n)]

    # --- Plan de financement
    avec_annee0 = (
        any(i.annee == 0 for i in entree.investissements)
        or any(a.annee == 0 for a in entree.apports)
        or any(p.annee_deblocage == 0 and p.inclure_plan_financement for p in entree.prets)
    )
    offset = 1 if avec_annee0 else 0
    nc = n + offset
    colonnes = (["Année 0"] if avec_annee0 else []) + list(entree.annees)

    def col(annee: int) -> int:
        # annee 0 -> colonne 0 ; annee k -> colonne k-1+offset ; si pas d'annee 0, l'annee 0 est rattachee a l'annee 1
        return max(annee, 1 - offset) - 1 + offset

    ressources: List[Ligne] = []
    caf_col = [None] * offset + caf
    ressources.append(Ligne("CAF", caf_col))
    for p in entree.prets:
        if not p.inclure_plan_financement:
            continue
        v = [0] * nc
        v[col(p.annee_deblocage)] += _r(p.montant)
        ressources.append(Ligne(f"Emprunt - {p.libelle}", v))
    for ap in entree.apports:
        v = [0] * nc
        v[col(ap.annee)] += _r(ap.montant)
        ressources.append(Ligne(ap.libelle, v))
    entrants = [sum((l.valeurs[c] or 0) for l in ressources) for c in range(nc)]

    emplois: List[Ligne] = []
    inv_v = [0] * nc
    for inv in entree.investissements:
        inv_v[col(inv.annee)] += _r(inv.montant)
    emplois.append(Ligne("Investissement", inv_v))
    emplois.append(Ligne("Remboursement emprunt", [0] * offset + [_r(v) for v in capital_rembourse_plan]))
    if any(aug_bfr):
        emplois.append(Ligne("Augmentation du BFR", [0] * offset + aug_bfr))
    sortants = [sum((l.valeurs[c] or 0) for l in emplois) for c in range(nc)]
    nets = [entrants[c] - sortants[c] for c in range(nc)]
    cumul, s = [], _r(entree.tresorerie_initiale)
    for v in nets:
        s += v
        cumul.append(s)

    plan = PlanFinancement(colonnes, ressources, entrants, emplois, sortants, nets, cumul)

    # --- Couverture de la dette
    fcf = [rn[a] + dotations[a] for a in range(n)]
    service = [_r(interets[a] + capital_rembourse_plan[a]) for a in range(n)]
    couverture = CouvertureDette(
        annees=list(entree.annees),
        resultat_net=rn,
        amortissements=dotations,
        free_cash_flow=fcf,
        frais_financiers=frais_fin,
        ratio_couverture=[_ratio(fcf[a], frais_fin[a]) for a in range(n)],
        service_dette=service,
        dscr=[_ratio(caf[a], service[a]) for a in range(n)],
    )

    indicateurs = {
        "croissance_ca": [None] + [_ratio(ca[a] - ca[a - 1], ca[a - 1]) for a in range(1, n)],
        "croissance_rn": [None] + [_ratio(rn[a] - rn[a - 1], abs(rn[a - 1])) for a in range(1, n)],
        "taux_marge_brute": [_ratio(marge_brute[a], ca[a]) for a in range(n)],
        "achats_sur_ca": [_ratio(par_rubrique["achats_matieres"][a], ca[a]) for a in range(n)],
        "va_sur_ca": [_ratio(va[a], ca[a]) for a in range(n)],
        "ebe_sur_ca": [_ratio(ebe[a], ca[a]) for a in range(n)],
        "rn_sur_ca": [_ratio(rn[a], ca[a]) for a in range(n)],
        # "charges rapportees au CA" au sens des dossiers BEFA : (CA - RN) / CA
        "charges_sur_ca": [_ratio(ca[a] - rn[a], ca[a]) for a in range(n)],
    }

    return Projection(entree, compte, plan, couverture, bfr, aug_bfr, echeanciers, indicateurs)
