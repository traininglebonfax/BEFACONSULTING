"""
Moteur de generation Excel.

Regle d'or : ce module clone le gabarit restructure (model/Business_Plan_Template.xlsx)
et n'ecrit JAMAIS que dans des cellules de saisie deja existantes. Il ne modifie,
n'insere ni ne supprime aucune formule, ligne, colonne ou feuille.

Cellules de saisie ecrites (voir scripts/restructure_template.py pour le detail
de la disposition du gabarit) :
  1_HYPOTHESES
    6-17    PRIX unitaire produits (A,B,D)
    18-29   VOLUME initial mensuel produits (A,B,D)
    30-31   Taux de croissance A2 / A3
    33-34   Capacite max usine / jours ouvres
    35-46   Cout variable unitaire produits (A,B,D)
    50-69   Charges fixes (A,B)
    71-80   Parametres financiers + taux evolution charges
    84-95   Saisonnalite M1-M12
  2_PREVISION_CA
    6-17, 22-33, 35-46   Libelles produits (colonne A uniquement -
                          toutes les formules sont deja en place)
  4_INVESTISSEMENTS
    6-20    Libelle / Montant / Duree investissements (A,B,C)
"""
import copy
import os
import shutil
from typing import Optional

import openpyxl
from openpyxl.comments import Comment

from .models import ProjetInput, MAX_PRODUITS, MAX_CHARGES, MAX_INVESTISSEMENTS

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE_PATH = os.path.join(BASE_DIR, "model", "Business_Plan_Template.xlsx")

AI_COMMENT_AUTHOR = "IA Business Plan"


def _hide_unused_rows(ws, start_row: int, used_count: int, total_capacity: int) -> None:
    for i in range(used_count, total_capacity):
        ws.row_dimensions[start_row + i].hidden = True


def generate_workbook(projet: ProjetInput, out_path: str, template_path: Optional[str] = None) -> str:
    """Clone le gabarit et ecrit les donnees du projet dans les cellules de saisie.

    Retourne le chemin du fichier .xlsx produit (non recalcule - voir engine.recalc).
    """
    projet.validate()

    template_path = template_path or TEMPLATE_PATH
    if not os.path.exists(template_path):
        raise FileNotFoundError(
            f"Gabarit introuvable: {template_path}. "
            "Lancer scripts/restructure_template.py avant d'utiliser le moteur."
        )

    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    shutil.copyfile(template_path, out_path)

    wb = openpyxl.load_workbook(out_path)

    # Capture des formules de reference AVANT ecriture, pour verification finale
    # (garde-fou : aucune formule ne doit avoir change apres nos ecritures).
    formulas_before = _snapshot_formulas(wb)

    ws1 = wb["1_HYPOTHESES"]
    ws2 = wb["2_PREVISION_CA"]
    ws4 = wb["4_INVESTISSEMENTS"]

    n_prod = len(projet.produits)
    n_charges = len(projet.charges_fixes)
    n_invest = len(projet.investissements)

    # --- Produits : prix (6-17), volume (18-29), cout variable (35-46) ---
    for i, p in enumerate(projet.produits):
        r_prix = 6 + i
        r_vol = 18 + i
        r_cout = 35 + i
        ws1.cell(row=r_prix, column=1, value=f"Prix de vente unitaire - {p.nom}")
        ws1.cell(row=r_prix, column=2, value=p.prix_unitaire)
        ws1.cell(row=r_vol, column=1, value=f"Volume initial mensuel - {p.nom}")
        ws1.cell(row=r_vol, column=2, value=p.volume_initial_mensuel)
        ws1.cell(row=r_vol, column=3, value=p.unite_volume)
        ws1.cell(row=r_cout, column=1, value=f"Cout variable unitaire - {p.nom}")
        ws1.cell(row=r_cout, column=2, value=p.cout_variable_unitaire)

        # 2_PREVISION_CA : libelles produits (colonne A des 3 blocs, memes index)
        ws2.cell(row=6 + i, column=1, value=p.nom)
        ws2.cell(row=22 + i, column=1, value=f"Vol. {p.nom} ({p.unite_volume})")
        ws2.cell(row=35 + i, column=1, value=f"CA {p.nom} (FCFA)")

    # --- Croissance ---
    ws1["B30"] = projet.taux_croissance_a2
    ws1["B31"] = projet.taux_croissance_a3

    # --- Production (capacite / jours ouvres) ---
    if projet.capacite_max_mensuelle is not None:
        ws1["B33"] = projet.capacite_max_mensuelle
    ws1["B34"] = projet.jours_ouvres_mensuel

    # --- Charges fixes (50-69) ---
    for i, c in enumerate(projet.charges_fixes):
        r = 50 + i
        ws1.cell(row=r, column=1, value=c.nom)
        ws1.cell(row=r, column=2, value=c.montant_mensuel)

    # --- Parametres financiers (71-80) ---
    ws1["B72"] = projet.part_fonds_propres
    ws1["B73"] = projet.part_emprunt
    ws1["B74"] = projet.taux_interet_annuel
    ws1["B75"] = projet.duree_remboursement_mois
    ws1["B76"] = projet.taux_is
    ws1["B77"] = projet.delai_clients_jours
    ws1["B78"] = projet.delai_fournisseurs_jours
    ws1["B79"] = projet.taux_evolution_charges_a2
    ws1["B80"] = projet.taux_evolution_charges_a3

    # --- Investissements (6-20) ---
    for i, inv in enumerate(projet.investissements):
        r = 6 + i
        ws4.cell(row=r, column=1, value=inv.nom)
        ws4.cell(row=r, column=2, value=inv.montant)
        ws4.cell(row=r, column=3, value=inv.duree_amortissement_ans)
    # Investissement total (B71 -> ligne 71 de 1_HYPOTHESES) : recopie de la
    # valeur informative, la vraie somme utilisee par les formules vient de
    # 4_INVESTISSEMENTS!B21 (SUM automatique).
    total_invest = sum(inv.montant for inv in projet.investissements)
    ws1["B71"] = total_invest

    # --- Saisonnalite (84-95) ---
    saison = projet.saisonnalite or [1.0] * 12
    for m in range(12):
        ws1.cell(row=84 + m, column=2, value=saison[m])

    # --- Masquage des lignes inutilisees pour un apercu propre ---
    _hide_unused_rows(ws1, 6, n_prod, MAX_PRODUITS)     # prix
    _hide_unused_rows(ws1, 18, n_prod, MAX_PRODUITS)    # volume
    _hide_unused_rows(ws1, 35, n_prod, MAX_PRODUITS)    # cout
    _hide_unused_rows(ws1, 50, n_charges, MAX_CHARGES)  # charges
    _hide_unused_rows(ws4, 6, n_invest, MAX_INVESTISSEMENTS)  # investissements
    _hide_unused_rows(ws2, 6, n_prod, MAX_PRODUITS)     # synthese produits
    _hide_unused_rows(ws2, 22, n_prod, MAX_PRODUITS)    # ventilation volume
    _hide_unused_rows(ws2, 35, n_prod, MAX_PRODUITS)    # ventilation CA

    # --- Garde-fou : verifier qu'aucune formule n'a change ---
    formulas_after = _snapshot_formulas(wb)
    diff = _diff_formulas(formulas_before, formulas_after)
    if diff:
        raise RuntimeError(
            "Le moteur a modifie des cellules contenant des formules - "
            f"violation de la regle d'or. Cellules concernees: {diff[:10]}"
        )

    wb.save(out_path)
    return out_path


def attach_ai_comments(xlsx_path: str, description_activite: str, justification_croissance: str) -> None:
    """Attache les textes narratifs generes par l'IA (Lot 4) en COMMENTAIRES
    de cellule, sur des cellules d'en-tete deja existantes de 1_HYPOTHESES.

    Ne touche a AUCUNE valeur de cellule ni formule (regle d'or) - un
    commentaire est une metadonnee separee de la valeur. Appele apres le
    recalcul LibreOffice (le recalcul ne connait pas les commentaires).
    """
    if not description_activite and not justification_croissance:
        return

    wb = openpyxl.load_workbook(xlsx_path)
    ws1 = wb["1_HYPOTHESES"]

    if description_activite:
        ws1["A1"].comment = Comment(
            f"Description de l'activite (generee par IA) :\n\n{description_activite}",
            AI_COMMENT_AUTHOR,
        )
    if justification_croissance:
        ws1["A30"].comment = Comment(
            f"Justification du profil de croissance (generee par IA) :\n\n{justification_croissance}",
            AI_COMMENT_AUTHOR,
        )

    wb.save(xlsx_path)


def attach_cost_suggestion_comments(xlsx_path: str, cost_suggestions: dict) -> None:
    """Attache un commentaire de cellule sur chaque cout variable unitaire
    (1_HYPOTHESES!B{35+i}) suggere par l'IA (role optionnel du Lot 4, cahier
    des charges Section2) faute de valeur fournie par l'utilisateur.

    `cost_suggestions` : {index_produit: valeur_suggeree}. Ne touche a AUCUNE
    formule - la valeur elle-meme est deja ecrite comme une saisie normale par
    generate_workbook ; ce commentaire sert uniquement a l'identifier comme
    suggestion IA (reste modifiable par l'utilisateur, cf. regle du cahier
    des charges) - voir aussi engine.preview pour le badge dans l'apercu web.
    """
    if not cost_suggestions:
        return

    wb = openpyxl.load_workbook(xlsx_path)
    ws1 = wb["1_HYPOTHESES"]

    for index, valeur in cost_suggestions.items():
        cell = ws1.cell(row=35 + index, column=2)
        cell.comment = Comment(
            f"Cout variable unitaire suggere par l'IA (valeur non fournie) : "
            f"{valeur:,.0f} FCFA. Modifiable - resaisissez le champ correspondant "
            "dans le formulaire et regenerez l'apercu si besoin.",
            AI_COMMENT_AUTHOR,
        )

    wb.save(xlsx_path)


def _snapshot_formulas(wb) -> dict:
    snap = {}
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and v.startswith("="):
                    snap[(sheet, cell.coordinate)] = v
    return snap


def _diff_formulas(before: dict, after: dict) -> list:
    changed = []
    for k, v in before.items():
        if after.get(k) != v:
            changed.append((k, v, after.get(k)))
    # une formule qui aurait disparu (ecrasee par une valeur) est aussi une violation
    return changed
