"""
Commentaires des tableaux financiers, rediges par le PROGRAMME (pas par
l'IA) dans le style des dossiers BEFA : aucun chiffre ne peut etre invente
ou mal recopie. Chaque fonction retourne une liste de paragraphes.
"""
from typing import List, Optional

from .historique import AnalyseHistorique
from .projection import Projection


def montant(x: float, devise: str = "FCFA") -> str:
    """16 084 640 000 -> '16,08 milliards FCFA' ; 182 466 271 -> '182 millions FCFA'."""
    signe = "−" if x < 0 else ""
    a = abs(x)
    if a >= 1e9:
        v = a / 1e9
        txt = f"{v:.2f}".rstrip("0").rstrip(".").replace(".", ",")
        unite = "milliard" if v < 2 else "milliards"
    elif a >= 1e6:
        v = a / 1e6
        txt = f"{v:.0f}" if v >= 100 else f"{v:.1f}".rstrip("0").rstrip(".").replace(".", ",")
        unite = "million" if v < 2 else "millions"
    else:
        return f"{signe}{entier(a)} {devise}"
    return f"{signe}{txt} {unite} {devise}"


def entier(x: float) -> str:
    return f"{x:,.0f}".replace(",", " ")


def pct(x: Optional[float], signe: bool = False) -> str:
    if x is None:
        return "n.d."
    s = f"{x * 100:+.0f}" if signe else f"{x * 100:.0f}"
    return s.replace("-", "−") + " %"


def tcam(debut: float, fin: float, annees: int) -> Optional[float]:
    if debut <= 0 or fin <= 0 or annees <= 0:
        return None
    return (fin / debut) ** (1 / annees) - 1


# ----------------------------------------------------------------------
# Libelles d'annees : "Année 3" (rang) ou "2026" (millesime)
# ----------------------------------------------------------------------
def _rang(l: str) -> bool:
    return l.strip().lower().startswith("année")


def en(l: str) -> str:
    return f"en {l.strip().lower()}" if _rang(l) else f"en {l}"


def des(l: str) -> str:
    return f"dès l'{l.strip().lower()}" if _rang(l) else f"dès {l}"


def a_partir(l: str) -> str:
    return f"à partir de l'{l.strip().lower()}" if _rang(l) else f"à partir de {l}"


def jusqu(l: str) -> str:
    return f"jusqu'à l'{l.strip().lower()}" if _rang(l) else f"jusqu'en {l}"


def fin_de(l: str) -> str:
    return f"à la fin de l'{l.strip().lower()}" if _rang(l) else f"fin {l}"


PERIODICITE = {"mensuelle": "mensuels", "trimestrielle": "trimestriels", "semestrielle": "semestriels",
               "annuelle": "annuels"}


def _premier_plateau(valeurs: List[float]) -> Optional[int]:
    """Indice a partir duquel la serie ne varie plus (ex. capacite nominale atteinte)."""
    for i in range(len(valeurs) - 1):
        if all(abs(v - valeurs[i]) <= 1e-6 * max(1, abs(valeurs[i])) for v in valeurs[i:]):
            return i
    return None


# ----------------------------------------------------------------------
# Previsions
# ----------------------------------------------------------------------
def compte_exploitation(p: Projection, nom: str) -> List[str]:
    ce = p.compte
    a, n = ce.annees, len(ce.annees)
    ca = ce.chiffre_affaires
    paras = []
    plateau = _premier_plateau(ca)
    if n == 1:
        paras.append(f"{nom} prévoit un chiffre d'affaires de {montant(ca[0])} {en(a[0])}.")
    elif plateau is not None and 0 < plateau < n - 1:
        paras.append(
            f"{nom} escompte un chiffre d'affaires qui s'améliore progressivement, passant de {montant(ca[0])} "
            f"{en(a[0])} à {montant(ca[plateau])} {a_partir(a[plateau])}, niveau maintenu {jusqu(a[-1])}."
        )
    else:
        t = tcam(ca[0], ca[-1], n - 1)
        paras.append(
            f"{nom} escompte un chiffre d'affaires passant de {montant(ca[0])} {en(a[0])} à {montant(ca[-1])} "
            f"{en(a[-1])}" + (f", soit une croissance moyenne de {pct(t)} par an." if t is not None else ".")
        )
    charges = [(ca[i] - ce.ebe[i]) / ca[i] if ca[i] else None for i in range(n)]
    if n > 1 and charges[0] is not None and charges[-1] is not None:
        ecart = (charges[-1] - charges[0]) * 100
        sens = ("restent stables" if abs(ecart) < 1 else
                f"perdent environ {abs(ecart):.0f} point{'s' if abs(ecart) >= 2 else ''}" if ecart < 0 else
                f"gagnent environ {ecart:.0f} point{'s' if ecart >= 2 else ''}")
        paras.append(
            f"Rapportées au chiffre d'affaires, les charges d'exploitation (hors amortissements) {sens}, "
            f"passant de {pct(charges[0])} {en(a[0])} à {pct(charges[-1])} {en(a[-1])}."
        )
    rn = ce.resultat_net
    if rn[0] < 0:
        positif = next((i for i, v in enumerate(rn) if v > 0), None)
        paras.append(
            f"Le résultat net est déficitaire {en(a[0])} ({montant(rn[0])})" +
            (f" et devient bénéficiaire {des(a[positif])} ({montant(rn[positif])}), pour atteindre "
             f"{montant(rn[-1])} {en(a[-1])}." if positif is not None else " et le reste sur toute la période.")
        )
    elif n > 1:
        t = tcam(rn[0], rn[-1], n - 1)
        paras.append(
            f"Le résultat net passe de {montant(rn[0])} {en(a[0])} à {montant(rn[-1])} {en(a[-1])}" +
            (f", soit une progression moyenne de {pct(t)} par an." if t is not None and t > 0 else ".")
        )
    return paras


def financement(p: Projection, nom: str) -> List[str]:
    total = p.cout_total_investissements
    paras = []
    if total:
        apports, prets = p.total_apports, p.total_prets_investissement
        paras.append(
            f"Le coût total des investissements s'élève à {entier(total)} FCFA. Il est financé par un apport du "
            f"promoteur de {entier(apports)} FCFA, soit {pct(apports / total)} du coût total"
            + (f", et par des concours bancaires de {entier(prets)} FCFA ({pct(prets / total)})." if prets else ".")
        )
    for pr in p.entree.prets:
        differe = []
        if pr.differe_total_mois:
            differe.append(f"{pr.differe_total_mois} mois de différé total")
        if pr.differe_partiel_mois:
            differe.append(f"{pr.differe_partiel_mois} mois de différé partiel")
        taux = f"{pr.taux_annuel * 100:.2f}".rstrip("0").rstrip(".").replace(".", ",")
        paras.append(
            f"{pr.libelle_type} « {pr.libelle} » : {entier(pr.montant)} FCFA au taux de {taux} % sur "
            f"{pr.duree_mois} mois" + (f" dont {' et '.join(differe)}" if differe else "") +
            (", remboursement in fine." if pr.in_fine else f", remboursements {PERIODICITE.get(pr.periodicite, pr.periodicite)}.")
        )
    return paras


def plan_financement(p: Projection, nom: str) -> List[str]:
    plan, ce = p.plan, p.compte
    a = ce.annees
    paras = [
        f"{nom} dégage une capacité d'autofinancement passant de {montant(ce.caf[0])} {en(a[0])} à "
        f"{montant(ce.caf[-1])} {en(a[-1])}."
    ]
    negatifs = [plan.colonnes[i] for i, v in enumerate(plan.flux_nets) if v < 0]
    if not negatifs:
        paras.append(
            f"Les flux nets de trésorerie restent positifs chaque année ; leur cumul atteint "
            f"{montant(plan.cumul[-1])} {fin_de(plan.colonnes[-1])}, ce qui traduit une capacité suffisante "
            f"à couvrir le service de la dette et les investissements prévus."
        )
    else:
        paras.append(
            f"Les flux nets de trésorerie sont négatifs {', '.join(en(x) for x in negatifs)}. Le cumul s'établit à "
            f"{montant(plan.cumul[-1])} en fin de période"
            + (" ; un besoin de trésorerie temporaire est à couvrir." if min(plan.cumul) < 0 else ".")
        )
    return paras


def couverture_dette(p: Projection) -> List[str]:
    cv = p.couverture
    a = cv.annees
    ratios = [(a[i], r) for i, r in enumerate(cv.ratio_couverture) if r is not None]
    if not ratios:
        return ["Le projet ne supporte pas de frais financiers sur la période : le ratio de couverture de la dette "
                "n'est pas applicable."]
    paras = [
        f"Le ratio de couverture de la dette (free cash-flow / frais financiers) évolue de "
        f"{ratios[0][1]:.1f}".replace(".", ",") + f" {en(ratios[0][0])} à " +
        f"{ratios[-1][1]:.1f}".replace(".", ",") + f" {en(ratios[-1][0])}."
    ]
    sous = [an for an, r in ratios if r < 1]
    if not sous:
        paras.append("Sur toute la période de remboursement, ce ratio est supérieur à 1 : l'entreprise pourra "
                      "faire face à ses engagements.")
    else:
        paras.append(f"Ce ratio est inférieur à 1 {', '.join(en(x) for x in sous)} : le free cash-flow ne couvre pas les frais "
                     f"financiers ces années-là, ce qui appelle un renforcement des fonds propres ou un "
                     f"réaménagement de la dette.")
    dscr = [(a[i], r) for i, r in enumerate(cv.dscr) if r is not None]
    if dscr:
        an_min, mini = min(dscr, key=lambda x: x[1])
        paras.append(
            f"Le taux de couverture du service de la dette (CAF / intérêts et capital remboursé) est au plus bas "
            f"{en(an_min)} à " + f"{mini:.2f}".replace(".", ",") +
            (" : la CAF couvre l'intégralité des échéances chaque année." if mini >= 1 else
             " : la CAF ne couvre pas entièrement les échéances cette année-là.")
        )
    return paras


# ----------------------------------------------------------------------
# Historique (entreprise existante)
# ----------------------------------------------------------------------
def analyse_bilan(h: AnalyseHistorique, nom: str) -> List[str]:
    n, p = h.bilan_n, h.bilan_n1
    r = n.ratios()
    paras = []
    if p is not None and p.actif_immobilise_net:
        var = (n.actif_immobilise_net - p.actif_immobilise_net) / p.actif_immobilise_net
        cause = ""
        if abs(n.actif_immobilise_brut - p.actif_immobilise_brut) < 1 and n.detail_amortissements and p.detail_amortissements:
            cause = (" Les immobilisations brutes étant inchangées, cette évolution s'explique par les amortissements : "
                     "l'entreprise n'a pas réalisé de nouvel investissement significatif sur l'exercice.")
        paras.append(f"Au terme de l'exercice {n.exercice}, les emplois stables évoluent de {pct(var, True)}.{cause}")
    if p is not None:
        paras.append(f"Les capitaux propres passent de {montant(p.capitaux_propres)} en {p.exercice} à "
                     f"{montant(n.capitaux_propres)} en {n.exercice}, pour un ratio de solvabilité (capitaux propres / "
                     f"total bilan) de {pct(r['solvabilite'])}.")
    else:
        paras.append(f"Les capitaux propres s'élèvent à {montant(n.capitaux_propres)} en {n.exercice}, pour un ratio "
                     f"de solvabilité de {pct(r['solvabilite'])}.")
    fr, bfr, tn = n.fonds_de_roulement, n.bfr, n.tresorerie_nette
    if fr >= bfr and fr > 0:
        paras.append(f"Le fonds de roulement ({montant(fr)}) couvre entièrement le besoin en fonds de roulement "
                     f"({montant(bfr)}) : l'exercice se clôture avec une trésorerie nette de {montant(tn)}.")
    else:
        paras.append(f"Le fonds de roulement ({montant(fr)}) ne couvre pas le besoin en fonds de roulement "
                     f"({montant(bfr)}) : la trésorerie nette ressort à {montant(tn)}.")
    if r["rentabilite_capitaux_propres"] is not None:
        paras.append(f"La rentabilité des capitaux propres (résultat net / capitaux propres) est de "
                     f"{pct(r['rentabilite_capitaux_propres'])}.")
    sain = n.capitaux_propres > 0 and tn >= 0 and (r["solvabilite"] or 0) >= 0.25
    paras.append(f"En conclusion, {nom} présente au {n.exercice} une situation financière "
                 + ("saine." if sain else "à surveiller."))
    return paras


def analyse_compte_resultat(h: AnalyseHistorique, nom: str) -> List[str]:
    n, p = h.cr_n, h.cr_n1
    paras = []
    if p is not None:
        paras.append(f"{nom} clôture l'exercice {n.exercice} avec un résultat net de {montant(n.resultat_net)}, "
                     f"contre {montant(p.resultat_net)} l'exercice précédent.")
        if p.chiffre_affaires:
            var = n.chiffre_affaires - p.chiffre_affaires
            paras.append(f"Le chiffre d'affaires s'établit à {montant(n.chiffre_affaires)}, en "
                         f"{'hausse' if var >= 0 else 'baisse'} de {montant(abs(var))} "
                         f"({pct(var / p.chiffre_affaires, True)}).")
        charges = lambda c: c.chiffre_affaires - c.resultat_exploitation
        if n.chiffre_affaires:
            paras.append(f"Les charges d'exploitation passent de {montant(charges(p))} à {montant(charges(n))}, soit "
                         f"{pct(charges(n) / n.chiffre_affaires)} du chiffre d'affaires de l'exercice.")
    else:
        paras.append(f"{nom} réalise sur l'exercice {n.exercice} un chiffre d'affaires de {montant(n.chiffre_affaires)} "
                     f"pour un résultat net de {montant(n.resultat_net)}.")
    return paras


def bilan_ouverture(b, nom: str) -> List[str]:
    """Commentaire du bilan d'ouverture PREVISIONNEL d'une startup."""
    total = b.total_passif
    if not total:
        return []
    fp = b.capitaux_propres
    paras = [f"Le bilan d'ouverture prévisionnel de {nom} totalise {montant(total)}. Les ressources de démarrage se "
             f"composent de fonds propres et quasi-fonds propres pour {montant(fp)} ({pct(fp / total)}), dont "
             f"{montant(b.capital)} de capital"
             + (f" et {montant(b.compte_courant_associes)} d'apports en compte courant d'associés"
                if b.compte_courant_associes else "")
             + (f", et d'emprunts pour {montant(b.dettes_financieres)} ({pct(b.dettes_financieres / total)})."
                if b.dettes_financieres else ".")]
    emplois = [f"{montant(b.actif_immobilise_net)} d'immobilisations"] if b.actif_immobilise_net else []
    if b.stocks:
        emplois.append(f"{montant(b.stocks)} de stocks initiaux")
    if b.tresorerie_actif:
        emplois.append(f"{montant(b.tresorerie_actif)} de disponibilités constituant le fonds de démarrage")
    if emplois:
        paras.append("Ces ressources financent " + ", ".join(emplois) + ".")
    if b.tresorerie_passif:
        paras.append(f"Les emplois excèdent les ressources de {montant(b.tresorerie_passif)} : ce besoin non couvert "
                     "doit être financé avant le démarrage.")
    return paras


def evolution(colonnes: List[str], lignes, nom: str) -> List[str]:
    """Commentaire de l'evolution sur trois exercices (TCAM du chiffre d'affaires et du resultat net)."""
    par = {l.libelle: l.valeurs for l in lignes}
    debut, fin = colonnes[0], colonnes[-2]
    paras = []
    ca = par.get("Chiffre d'affaires")
    if ca and ca[0] and ca[-2] is not None:
        t = ca[-1]
        paras.append(f"Sur la période {debut}-{fin}, le chiffre d'affaires de {nom} passe de {montant(ca[0])} à "
                     f"{montant(ca[-2])}" + (f", soit une croissance annuelle moyenne de {pct(t)}." if t is not None else "."))
    rn = par.get("Résultat net")
    if rn and rn[-2] is not None:
        positifs = sum(1 for v in rn[:-1] if v is not None and v > 0)
        paras.append(f"Le résultat net est bénéficiaire sur {positifs} des {len(rn) - 1} exercices analysés et "
                     f"s'établit à {montant(rn[-2])} en {fin}.")
    cp = par.get("Capitaux propres")
    if cp and cp[0] is not None and cp[-2] is not None:
        paras.append(f"Les capitaux propres évoluent de {montant(cp[0])} à {montant(cp[-2])} sur la période.")
    return paras
