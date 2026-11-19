"""
Bilan d'ouverture PREVISIONNEL d'une startup (aucun historique comptable).

Proposition automatique, calculee par le programme a partir des autres
etapes du dossier (jamais par l'IA, pour garantir l'equilibre et l'absence
de chiffre invente) :

  ACTIF  : investissements de la phase de demarrage (annee 0) -> immobilisations ;
           stocks initiaux et creances eventuels (saisis) ;
           disponibilites = solde des ressources non employees.
  PASSIF : capital social envisage ; apports des promoteurs au-dela du capital
           -> compte courant d'associes ; subventions ; emprunts debloques en
           annee 0 ; dettes initiales eventuelles (fournisseurs).

Si les emplois depassent les ressources, l'ecart apparait en tresorerie
passive (besoin de financement non couvert) et une alerte est emise.
Le consultant modifie les montants puis valide le bilan. Ce n'est jamais une
donnee comptable historique : le document le presente comme une estimation.
"""
from dataclasses import dataclass, field
from typing import List, Optional

from .historique import BilanCondense, PosteImmobilisation
from .projection import ProjectionInput

EXERCICE = "Ouverture (prévisionnel)"
MENTION = ("Bilan d'ouverture prévisionnel établi à partir des hypothèses du projet (investissements, apports, "
           "emprunts sollicités) : il s'agit d'une estimation et non d'un bilan comptable certifié.")


@dataclass
class PropositionOuverture:
    bilan: BilanCondense
    hypotheses: List[str] = field(default_factory=list)
    alertes: List[str] = field(default_factory=list)


def _f(x: float) -> str:
    return f"{x:,.0f}".replace(",", " ")


def proposer(projection: Optional[ProjectionInput], capital_envisage: Optional[float] = None,
             stocks_initiaux: float = 0, creances_initiales: float = 0, dettes_initiales: float = 0) -> PropositionOuverture:
    hyp: List[str] = []
    alertes: List[str] = []
    b = BilanCondense(exercice=EXERCICE)
    if projection is None:
        alertes.append("Aucune donnée financière : renseignez les investissements, apports et emprunts (étape "
                       "« Données financières ») pour obtenir une proposition complète.")
        projection = ProjectionInput(annees=["Année 1"], produits=[])

    # Actif immobilise : investissements de la phase de demarrage
    groupes: dict = {}
    for inv in projection.investissements:
        if inv.annee == 0 and inv.montant:
            cle = inv.rubrique or inv.libelle
            groupes[cle] = groupes.get(cle, 0) + float(inv.montant)
    b.immobilisations = [PosteImmobilisation(lib, round(m), 0) for lib, m in groupes.items()]
    if groupes:
        hyp.append(f"Immobilisations : investissements prévus avant le démarrage ({_f(sum(groupes.values()))} FCFA), "
                   "non encore amortis à l'ouverture.")
    tardifs = [i for i in projection.investissements if i.annee > 0]
    if tardifs:
        hyp.append(f"{len(tardifs)} investissement(s) prévu(s) après l'ouverture exclu(s) du bilan d'ouverture.")

    b.stocks = float(stocks_initiaux or 0)
    b.creances_clients = float(creances_initiales or 0)
    b.fournisseurs = float(dettes_initiales or 0)
    if b.stocks:
        hyp.append(f"Stocks initiaux saisis : {_f(b.stocks)} FCFA.")

    # Ressources
    apports = [a for a in projection.apports if a.annee == 0]
    subv = sum(a.montant for a in apports if a.nature == "subvention")
    cca = sum(a.montant for a in apports if a.nature in ("compte courant", "compte_courant"))
    fonds_propres = sum(a.montant for a in apports if a.nature not in ("subvention", "compte courant", "compte_courant"))
    if capital_envisage:
        b.capital = float(capital_envisage)
        surplus = fonds_propres - b.capital
        if surplus > 0:
            cca += surplus
            hyp.append(f"Capital social envisagé : {_f(b.capital)} FCFA ; le surplus des apports des promoteurs "
                       f"({_f(surplus)} FCFA) est porté en compte courant d'associés.")
        elif surplus < 0:
            alertes.append(f"Les apports des promoteurs ({_f(fonds_propres)} FCFA) sont inférieurs au capital envisagé "
                           f"({_f(b.capital)} FCFA) : capital non entièrement libéré ou apports à compléter.")
            b.autres_creances += -surplus  # capital souscrit non verse (creance sur associes)
            hyp.append(f"Capital souscrit non versé : {_f(-surplus)} FCFA (créance sur les associés).")
    else:
        b.capital = float(fonds_propres)
        if fonds_propres:
            hyp.append(f"Capital : apports des promoteurs ({_f(fonds_propres)} FCFA), en l'absence de capital "
                       "social renseigné.")
    b.compte_courant_associes = float(cca)
    b.subventions_investissement = float(subv)
    if subv:
        hyp.append(f"Subventions d'investissement : {_f(subv)} FCFA.")
    prets = [p for p in projection.prets if p.annee_deblocage == 0]
    b.dettes_financieres = float(sum(p.montant for p in prets))
    if prets:
        hyp.append(f"Emprunts débloqués à l'ouverture : {_f(b.dettes_financieres)} FCFA "
                   f"({', '.join(p.libelle or p.type for p in prets)}).")

    # Equilibre : disponibilites (ou besoin non couvert)
    emplois = b.actif_immobilise_net + b.actif_circulant
    ressources = b.capitaux_permanents + b.passif_circulant
    solde = ressources - emplois
    if solde >= 0:
        b.tresorerie_actif = float(solde)
        if solde:
            hyp.append(f"Disponibilités : ressources non employées en investissements ({_f(solde)} FCFA), "
                       "constituant le fonds de démarrage.")
    else:
        b.tresorerie_passif = float(-solde)
        alertes.append(f"Les emplois dépassent les ressources de {_f(-solde)} FCFA : besoin de financement non couvert "
                       "(porté en trésorerie passive). Augmentez les apports ou le financement sollicité.")
    if not (b.immobilisations or b.capital or b.dettes_financieres):
        alertes.append("Bilan d'ouverture vide : aucun investissement, apport ni emprunt de démarrage renseigné.")
    return PropositionOuverture(b, hyp, alertes)
