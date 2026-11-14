"""
Analyse financiere historique (variante "entreprise existante", modele
O'IVOIRE) : bilan condense et compte de resultat condense sur deux
exercices, avec variations N / N-1 et ratios de structure.

Les donnees viennent des liasses SYSCOHADA importees (Lot 2), toujours
validees par le consultant avant usage.
"""
from dataclasses import dataclass, field
from typing import Dict, List, Optional

from .projection import Ligne


def _pct(a: float, b: float) -> Optional[float]:
    return None if not b else a / b


@dataclass
class PosteImmobilisation:
    libelle: str
    brut: float = 0
    amortissements: float = 0

    @property
    def net(self) -> float:
        return self.brut - self.amortissements


@dataclass
class BilanCondense:
    exercice: str
    immobilisations: List[PosteImmobilisation] = field(default_factory=list)
    # Actif circulant
    stocks: float = 0
    avances_fournisseurs: float = 0
    creances_clients: float = 0
    autres_creances: float = 0
    tresorerie_actif: float = 0
    # Capitaux propres (BEFA y inclut le compte courant d'associes)
    capital: float = 0
    reserves_report: float = 0
    resultat_net: float = 0
    subventions_investissement: float = 0
    compte_courant_associes: float = 0
    # Dettes
    dettes_financieres: float = 0
    fournisseurs: float = 0
    dettes_fiscales: float = 0
    dettes_sociales: float = 0
    dettes_fiscales_sociales: float = 0  # ligne DK de la liasse SYSCOHADA (non ventilee)
    autres_dettes: float = 0
    tresorerie_passif: float = 0
    # False quand seule la valeur nette des immobilisations est connue
    # (colonne N-1 d'une liasse : brut et amortissements non fournis)
    detail_amortissements: bool = True

    @property
    def actif_immobilise_brut(self) -> float:
        return sum(p.brut for p in self.immobilisations)

    @property
    def actif_immobilise_net(self) -> float:
        return sum(p.net for p in self.immobilisations)

    @property
    def actif_circulant(self) -> float:
        return self.stocks + self.avances_fournisseurs + self.creances_clients + self.autres_creances

    @property
    def total_actif(self) -> float:
        return self.actif_immobilise_net + self.actif_circulant + self.tresorerie_actif

    @property
    def capitaux_propres(self) -> float:
        return (self.capital + self.reserves_report + self.resultat_net
                + self.subventions_investissement + self.compte_courant_associes)

    @property
    def capitaux_permanents(self) -> float:
        return self.capitaux_propres + self.dettes_financieres

    @property
    def passif_circulant(self) -> float:
        return (self.fournisseurs + self.dettes_fiscales + self.dettes_sociales
                + self.dettes_fiscales_sociales + self.autres_dettes)

    @property
    def total_passif(self) -> float:
        return self.capitaux_permanents + self.passif_circulant + self.tresorerie_passif

    @property
    def ecart_equilibre(self) -> float:
        """Actif - passif : doit etre nul pour un bilan correctement saisi."""
        return self.total_actif - self.total_passif

    @property
    def fonds_de_roulement(self) -> float:
        return self.capitaux_permanents - self.actif_immobilise_net

    @property
    def bfr(self) -> float:
        return self.actif_circulant - self.passif_circulant

    @property
    def tresorerie_nette(self) -> float:
        return self.tresorerie_actif - self.tresorerie_passif

    def ratios(self) -> Dict[str, Optional[float]]:
        return {
            "solvabilite": _pct(self.capitaux_propres, self.total_passif),  # CP / total bilan
            "autonomie_financiere": _pct(self.capitaux_propres, self.capitaux_propres + self.dettes_financieres),
            "rentabilite_capitaux_propres": _pct(self.resultat_net, self.capitaux_propres),
            "liquidite_generale": _pct(self.actif_circulant + self.tresorerie_actif,
                                       self.passif_circulant + self.tresorerie_passif),
            "endettement": _pct(self.dettes_financieres, self.capitaux_propres),
        }

    @staticmethod
    def from_dict(d: dict) -> "BilanCondense":
        k = dict(d)
        k["immobilisations"] = [PosteImmobilisation(**x) for x in d.get("immobilisations", [])]
        return BilanCondense(**k)


@dataclass
class CompteResultatCondense:
    exercice: str
    chiffre_affaires: float = 0
    achats_matieres: float = 0  # achats de marchandises / matieres, variation de stocks incluse
    autres_achats: float = 0
    transport: float = 0
    services_exterieurs: float = 0
    impots_taxes: float = 0
    autres_charges: float = 0  # nettes des autres produits d'exploitation
    salaires: float = 0
    dotations: float = 0
    frais_financiers: float = 0  # nets des produits financiers
    resultat_hao: float = 0  # hors activites ordinaires (produits - charges)
    impot_sur_benefice: float = 0

    @property
    def marge_brute(self) -> float:
        return self.chiffre_affaires - self.achats_matieres

    @property
    def valeur_ajoutee(self) -> float:
        return (self.marge_brute - self.autres_achats - self.transport
                - self.services_exterieurs - self.impots_taxes - self.autres_charges)

    @property
    def ebe(self) -> float:
        return self.valeur_ajoutee - self.salaires

    @property
    def resultat_exploitation(self) -> float:
        return self.ebe - self.dotations

    @property
    def resultat_net(self) -> float:
        return self.resultat_exploitation - self.frais_financiers + self.resultat_hao - self.impot_sur_benefice

    @property
    def caf(self) -> float:
        return self.resultat_net + self.dotations


def _ligne_variation(libelle: str, n: float, n1: float, style: str = "") -> Ligne:
    var = n - n1
    return Ligne(libelle, [n, n1, var, _pct(var, n1)], style)


@dataclass
class AnalyseHistorique:
    bilan_n: BilanCondense
    bilan_n1: Optional[BilanCondense]
    cr_n: CompteResultatCondense
    cr_n1: Optional[CompteResultatCondense]

    @property
    def colonnes(self) -> List[str]:
        n1 = self.bilan_n1.exercice if self.bilan_n1 else "N-1"
        return [self.bilan_n.exercice, n1, "Var N/N-1 (val)", "Var N/N-1 (%)"]

    def _bn1(self) -> BilanCondense:
        return self.bilan_n1 or BilanCondense(exercice="N-1")

    def _crn1(self) -> CompteResultatCondense:
        return self.cr_n1 or CompteResultatCondense(exercice="N-1")

    def lignes_bilan(self) -> List[Ligne]:
        n, p = self.bilan_n, self._bn1()
        out: List[Ligne] = []
        libelles = [x.libelle for x in n.immobilisations]
        libelles += [x.libelle for x in p.immobilisations if x.libelle not in libelles]
        detail = n.detail_amortissements and p.detail_amortissements
        for lib in libelles:
            a = next((x for x in n.immobilisations if x.libelle == lib), PosteImmobilisation(lib))
            b = next((x for x in p.immobilisations if x.libelle == lib), PosteImmobilisation(lib))
            if not detail:
                # comparaison en valeur nette : le brut N-1 n'est pas connu
                out.append(_ligne_variation(f"{lib} (net)", a.net, b.net, "detail"))
                continue
            out.append(_ligne_variation(lib, a.brut, b.brut, "detail"))
            if a.amortissements or b.amortissements:
                out.append(_ligne_variation(f"Amortissement {lib}", -a.amortissements, -b.amortissements, "detail"))
        out.append(_ligne_variation("Emplois stables", n.actif_immobilise_net, p.actif_immobilise_net, "sous_total"))
        for lib, attr in (("Stocks", "stocks"), ("Fournisseurs, avances versées", "avances_fournisseurs"),
                          ("Créances clients", "creances_clients"), ("Autres créances", "autres_creances")):
            if getattr(n, attr) or getattr(p, attr):
                out.append(_ligne_variation(lib, getattr(n, attr), getattr(p, attr), "detail"))
        out.append(_ligne_variation("Actifs circulants", n.actif_circulant, p.actif_circulant, "sous_total"))
        out.append(_ligne_variation("Trésorerie", n.tresorerie_actif, p.tresorerie_actif, "sous_total"))
        out.append(_ligne_variation("Total Actif", n.total_actif, p.total_actif, "total"))
        for lib, attr in (("Capital", "capital"), ("Réserves et report à nouveau", "reserves_report"),
                          ("Résultat net", "resultat_net"), ("Subventions d'investissement", "subventions_investissement"),
                          ("Compte courant d'associés", "compte_courant_associes")):
            if getattr(n, attr) or getattr(p, attr):
                out.append(_ligne_variation(lib, getattr(n, attr), getattr(p, attr), "detail"))
        out.append(_ligne_variation("Capitaux propres", n.capitaux_propres, p.capitaux_propres, "sous_total"))
        if n.dettes_financieres or p.dettes_financieres:
            out.append(_ligne_variation("Dettes financières", n.dettes_financieres, p.dettes_financieres, "detail"))
        out.append(_ligne_variation("Capitaux permanents", n.capitaux_permanents, p.capitaux_permanents, "sous_total"))
        for lib, attr in (("Fournisseurs", "fournisseurs"), ("Dettes fiscales", "dettes_fiscales"),
                          ("Dettes sociales", "dettes_sociales"),
                          ("Dettes fiscales et sociales", "dettes_fiscales_sociales"), ("Autres dettes", "autres_dettes")):
            if getattr(n, attr) or getattr(p, attr):
                out.append(_ligne_variation(lib, getattr(n, attr), getattr(p, attr), "detail"))
        out.append(_ligne_variation("Passifs circulants", n.passif_circulant, p.passif_circulant, "sous_total"))
        if n.tresorerie_passif or p.tresorerie_passif:
            out.append(_ligne_variation("Trésorerie passif", n.tresorerie_passif, p.tresorerie_passif, "sous_total"))
        out.append(_ligne_variation("Total Passif", n.total_passif, p.total_passif, "total"))
        return out

    def lignes_structure(self) -> List[Ligne]:
        n, p = self.bilan_n, self._bn1()
        return [
            _ligne_variation("Fonds de roulement", n.fonds_de_roulement, p.fonds_de_roulement),
            _ligne_variation("Besoin en fonds de roulement", n.bfr, p.bfr),
            _ligne_variation("Trésorerie nette", n.tresorerie_nette, p.tresorerie_nette, "total"),
        ]

    def lignes_compte_resultat(self) -> List[Ligne]:
        n, p = self.cr_n, self._crn1()
        out = [_ligne_variation("Chiffre d'affaires", n.chiffre_affaires, p.chiffre_affaires)]
        if n.achats_matieres or p.achats_matieres:
            out.append(_ligne_variation("Achats de marchandises et matières", n.achats_matieres, p.achats_matieres))
        out.append(_ligne_variation("Marge brute", n.marge_brute, p.marge_brute, "sous_total"))
        for lib, attr in (("Autres achats", "autres_achats"), ("Transport", "transport"),
                          ("Services extérieurs", "services_exterieurs"), ("Impôts et taxes", "impots_taxes"),
                          ("Autres charges", "autres_charges")):
            if attr != "autres_charges" or n.autres_charges or p.autres_charges:
                out.append(_ligne_variation(lib, getattr(n, attr), getattr(p, attr)))
        out += [
            _ligne_variation("Valeur ajoutée", n.valeur_ajoutee, p.valeur_ajoutee, "sous_total"),
            _ligne_variation("Charges salariales", n.salaires, p.salaires),
            _ligne_variation("Excédent Brut d'Exploitation", n.ebe, p.ebe, "sous_total"),
            _ligne_variation("Dotations aux amortissements", n.dotations, p.dotations),
            _ligne_variation("Résultat d'exploitation", n.resultat_exploitation, p.resultat_exploitation, "sous_total"),
        ]
        if n.frais_financiers or p.frais_financiers:
            out.append(_ligne_variation("Frais financiers", n.frais_financiers, p.frais_financiers))
        if n.resultat_hao or p.resultat_hao:
            out.append(_ligne_variation("Résultat HAO", n.resultat_hao, p.resultat_hao))
        out += [
            _ligne_variation("Impôt sur les sociétés", n.impot_sur_benefice, p.impot_sur_benefice),
            _ligne_variation("Résultat net", n.resultat_net, p.resultat_net, "total"),
            _ligne_variation("CAF", n.caf, p.caf, "total"),
        ]
        return out

    def controles(self) -> List[str]:
        """Incoherences a signaler au consultant avant generation du PDF."""
        alertes = []
        for b in (self.bilan_n, self.bilan_n1):
            if b and abs(b.ecart_equilibre) > 1:
                alertes.append(f"Bilan {b.exercice} déséquilibré : actif - passif = {b.ecart_equilibre:,.0f} FCFA.".replace(",", " "))
        for b, cr in ((self.bilan_n, self.cr_n), (self.bilan_n1, self.cr_n1)):
            if b and cr and abs(b.resultat_net - cr.resultat_net) > 1:
                alertes.append(
                    f"Exercice {b.exercice} : le résultat net du bilan ({b.resultat_net:,.0f}) diffère de celui du "
                    f"compte de résultat ({cr.resultat_net:,.0f}).".replace(",", " ")
                )
        return alertes
