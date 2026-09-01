"""
Restructuration ponctuelle du classeur modele Business Plan.

Objectif : remplacer les blocs a taille fixe (3 produits, 8 charges fixes,
7 investissements) par des blocs provisionnes a taille max (12 produits,
20 charges, 15 investissements), sans changer un seul resultat calcule
pour le jeu de donnees "Boissons Naturelles" d'origine.

Regle : aucune formule n'est modifiee "a l'execution" par le moteur plus
tard - tout le cablage est fixe ici, une fois, sous forme versionnee.

Usage: python restructure_template.py
Entree : ../Business_Plan_Boissons_Naturelles_CI_avec_Guide_1.xlsx
Sortie : ../model/Business_Plan_Template.xlsx
"""
import copy
import os
import openpyxl
from openpyxl.utils import get_column_letter

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(BASE_DIR, "Business_Plan_Boissons_Naturelles_CI_avec_Guide_1.xlsx")
DST = os.path.join(BASE_DIR, "model", "Business_Plan_Template.xlsx")

N_PRODUITS = 12
N_CHARGES = 20
N_INVEST = 15
N_PRODUITS_ACTUELS = 3
N_CHARGES_ACTUELLES = 8
N_INVEST_ACTUELS = 7

wb = openpyxl.load_workbook(SRC)


def style_of(ws, coord):
    c = ws[coord]
    return dict(
        font=copy.copy(c.font),
        fill=copy.copy(c.fill),
        border=copy.copy(c.border),
        alignment=copy.copy(c.alignment),
        number_format=c.number_format,
    )


def apply_style(cell, style):
    cell.font = style["font"]
    cell.fill = style["fill"]
    cell.border = style["border"]
    cell.alignment = style["alignment"]
    cell.number_format = style["number_format"]


def row_vals(ws, r, ncols):
    return [ws.cell(row=r, column=col).value for col in range(1, ncols + 1)]


def set_row(ws, r, vals, styles):
    """styles: list of style-dict per column (1-indexed same length as vals)"""
    for i, v in enumerate(vals, start=1):
        cell = ws.cell(row=r, column=i)
        cell.value = v
        if styles[i - 1] is not None:
            apply_style(cell, styles[i - 1])


# =====================================================================
# 1) CAPTURE des donnees et styles de l'original avant toute mutation
# =====================================================================
ws1 = wb["1_HYPOTHESES"]

prix_rows = [row_vals(ws1, r, 4) for r in (6, 7, 8)]
vol_rows = [row_vals(ws1, r, 4) for r in (9, 10, 11)]
cout_mp_rows = [row_vals(ws1, r, 4) for r in (17, 18, 19)]
charge_rows = [row_vals(ws1, r, 4) for r in range(23, 31)]
season_rows = [row_vals(ws1, r, 3) for r in range(43, 55)]

single = {
    "croissance_a2": row_vals(ws1, 12, 4),
    "croissance_a3": row_vals(ws1, 13, 4),
    "capacite_max": row_vals(ws1, 15, 4),
    "jours_ouvres": row_vals(ws1, 16, 4),
    "packaging": row_vals(ws1, 20, 4),
    "transformation": row_vals(ws1, 21, 4),
    "capex_total": row_vals(ws1, 32, 4),
    "part_fp": row_vals(ws1, 33, 4),
    "part_emprunt": row_vals(ws1, 34, 4),
    "taux_interet": row_vals(ws1, 35, 4),
    "duree_remb": row_vals(ws1, 36, 4),
    "taux_is": row_vals(ws1, 37, 4),
    "delai_clients": row_vals(ws1, 38, 4),
    "delai_fournisseurs": row_vals(ws1, 39, 4),
}

sect_prod_couts = ws1["A14"].value
sect_opex = ws1["A22"].value
sect_fin = ws1["A31"].value
sect_saison = ws1["A41"].value
header_saison = row_vals(ws1, 42, 3)
header_table = row_vals(ws1, 4, 4)
title1 = ws1["A1"].value
subtitle1 = ws1["A2"].value

style_header = [style_of(ws1, c) for c in ("A4", "B4", "C4", "D4")]
style_section = style_of(ws1, "A5")
style_section_saison = style_of(ws1, "A41")
style_label = style_of(ws1, "A6")
style_input = style_of(ws1, "B6")
style_unit = style_of(ws1, "C6")
style_comment = style_of(ws1, "D6")
style_season_header = [style_of(ws1, c) for c in ("A42", "B42", "C42")]
style_season_label = style_of(ws1, "A43")
style_season_input = style_of(ws1, "B43")
style_season_comment = style_of(ws1, "C43")

ws2 = wb["2_PREVISION_CA"]
synth_rows = [row_vals(ws2, r, 8) for r in (6, 7, 8)]
synth_total_row = row_vals(ws2, 9, 8)
vent_header = row_vals(ws2, 12, 14)
vol_vent_rows = [row_vals(ws2, r, 14) for r in (13, 14, 15)]
vol_total_row = row_vals(ws2, 16, 14)
ca_vent_rows = [row_vals(ws2, r, 14) for r in (17, 18, 19)]
ca_total_row = row_vals(ws2, 20, 14)
title2 = ws2["A1"].value
subtitle2 = ws2["A2"].value
titre_synth = ws2["A4"].value
titre_vent = ws2["A11"].value

style2_header = [style_of(ws2, c) for c in ("A5", "B5", "C5", "D5", "E5", "F5", "G5", "H5")]
style2_prod_label = style_of(ws2, "A6")
style2_num = style_of(ws2, "B6")
style2_total = [style_of(ws2, c) for c in ("A9", "B9", "C9", "D9", "E9", "F9", "G9", "H9")]
style2_month_header = style_of(ws2, "A12")
style2_vent_num = style_of(ws2, "B13")
style2_vent_total = style_of(ws2, "N13")
style2_vol_total_row = style_of(ws2, "A16")
style2_ca_mensuel_total = style_of(ws2, "A20")

ws4 = wb["4_INVESTISSEMENTS"]
invest_rows = [row_vals(ws4, r, 6) for r in range(6, 13)]
title4 = ws4["A1"].value
subtitle4 = ws4["A2"].value
titre_capex = ws4["A4"].value
titre_financement = ws4["A15"].value
style4_header = [style_of(ws4, c) for c in ("A5", "B5", "C5", "D5", "E5", "F5")]
style4_label = style_of(ws4, "A6")
style4_num = style_of(ws4, "B6")
style4_total = [style_of(ws4, c) for c in ("A13", "B13", "C13", "D13", "E13", "F13")]
style4_fin_header = [style_of(ws4, c) for c in ("A16", "B16", "C16", "D16")]
row17 = row_vals(ws4, 17, 4)
row18 = row_vals(ws4, 18, 4)
row19 = row_vals(ws4, 19, 4)
style4_fin_line = [style_of(ws4, c) for c in ("A17", "B17", "C17", "D17")]
style4_fin_total = [style_of(ws4, c) for c in ("A19", "B19", "C19", "D19")]

print("Capture terminee.")

# =====================================================================
# 2) RECONSTRUCTION DE 1_HYPOTHESES
# =====================================================================
# Nouvelle disposition (voir plan valide) :
#   6-17   : PRIX unitaire produits x12
#   18-29  : VOLUME initial mensuel produits x12
#   30-31  : taux croissance A2 / A3
#   32     : section production & couts
#   33-34  : capacite max / jours ouvres
#   35-46  : COUT MP produits x12
#   47-48  : packaging / transformation
#   49     : section OPEX
#   50-69  : charges OPEX x20
#   70     : section financieres
#   71-78  : parametres financiers (8 lignes)
#   79-80  : NOUVEAU - taux evolution charges fixes A2 / A3
#   82     : section saisonnalite
#   83     : header saisonnalite
#   84-95  : M1-M12

for rng in list(ws1.merged_cells.ranges):
    ws1.unmerge_cells(str(rng))

for row in ws1.iter_rows(min_row=1, max_row=ws1.max_row, max_col=6):
    for cell in row:
        cell.value = None

ws1["A1"] = title1
ws1["A2"] = subtitle1
set_row(ws1, 4, header_table, style_header)
ws1.cell(row=5, column=1, value="HYPOTHESES COMMERCIALES & PRIX")
apply_style(ws1.cell(row=5, column=1), style_section)

for i in range(N_PRODUITS):
    r = 6 + i
    vals = prix_rows[i] if i < N_PRODUITS_ACTUELS else ["", None, "FCFA", ""]
    set_row(ws1, r, vals, [style_label, style_input, style_unit, style_comment])

for i in range(N_PRODUITS):
    r = 18 + i
    vals = vol_rows[i] if i < N_PRODUITS_ACTUELS else ["", None, "Unites/mois", ""]
    set_row(ws1, r, vals, [style_label, style_input, style_unit, style_comment])

set_row(ws1, 30, single["croissance_a2"], [style_label, style_input, style_unit, style_comment])
set_row(ws1, 31, single["croissance_a3"], [style_label, style_input, style_unit, style_comment])

ws1.cell(row=32, column=1, value=sect_prod_couts)
apply_style(ws1.cell(row=32, column=1), style_section)
set_row(ws1, 33, single["capacite_max"], [style_label, style_input, style_unit, style_comment])
set_row(ws1, 34, single["jours_ouvres"], [style_label, style_input, style_unit, style_comment])

for i in range(N_PRODUITS):
    r = 35 + i
    vals = cout_mp_rows[i] if i < N_PRODUITS_ACTUELS else ["", None, "FCFA/unite", ""]
    set_row(ws1, r, vals, [style_label, style_input, style_unit, style_comment])

set_row(ws1, 47, single["packaging"], [style_label, style_input, style_unit, style_comment])
set_row(ws1, 48, single["transformation"], [style_label, style_input, style_unit, style_comment])

ws1.cell(row=49, column=1, value=sect_opex)
apply_style(ws1.cell(row=49, column=1), style_section)

for i in range(N_CHARGES):
    r = 50 + i
    vals = charge_rows[i] if i < N_CHARGES_ACTUELLES else ["", None, "FCFA/mois", ""]
    set_row(ws1, r, vals, [style_label, style_input, style_unit, style_comment])

ws1.cell(row=70, column=1, value=sect_fin)
apply_style(ws1.cell(row=70, column=1), style_section)
set_row(ws1, 71, single["capex_total"], [style_label, style_input, style_unit, style_comment])
set_row(ws1, 72, single["part_fp"], [style_label, style_input, style_unit, style_comment])
set_row(ws1, 73, single["part_emprunt"], [style_label, style_input, style_unit, style_comment])
set_row(ws1, 74, single["taux_interet"], [style_label, style_input, style_unit, style_comment])
set_row(ws1, 75, single["duree_remb"], [style_label, style_input, style_unit, style_comment])
set_row(ws1, 76, single["taux_is"], [style_label, style_input, style_unit, style_comment])
set_row(ws1, 77, single["delai_clients"], [style_label, style_input, style_unit, style_comment])
set_row(ws1, 78, single["delai_fournisseurs"], [style_label, style_input, style_unit, style_comment])

# Nouveaux parametres generiques - taux d'evolution des charges fixes (remplacent les
# taux differencies par categorie nommee, decision validee avec l'utilisateur)
set_row(
    ws1, 79,
    ["Taux d'evolution des charges fixes Annee 2", 0.05, "%", "Evolution generique appliquee a la masse totale des charges fixes"],
    [style_label, style_input, style_unit, style_comment],
)
set_row(
    ws1, 80,
    ["Taux d'evolution des charges fixes Annee 3", 0.10, "%", "Evolution generique appliquee a la masse totale des charges fixes"],
    [style_label, style_input, style_unit, style_comment],
)

ws1.cell(row=82, column=1, value=sect_saison)
apply_style(ws1.cell(row=82, column=1), style_section_saison)
set_row(ws1, 83, header_saison, style_season_header)
for i in range(12):
    r = 84 + i
    set_row(ws1, r, season_rows[i], [style_season_label, style_season_input, style_season_comment])

# Formats de nombre corrects par ligne (perdus lors du style_input uniforme
# applique plus haut, qui reprenait le format FCFA de la ligne prix pour
# toutes les lignes - regression corrigee ici en reappliquant le format
# d'origine constate sur le classeur source, ligne par ligne).
FMT_FCFA = '#,##0" FCFA"'
FMT_NUM = '#,##0'
FMT_PCT = '0.0%'
FMT_DEC = '#,##0.00'

number_formats = {}
for i in range(N_PRODUITS):
    number_formats[18 + i] = FMT_NUM  # volume
number_formats[30] = FMT_PCT
number_formats[31] = FMT_PCT
number_formats[33] = FMT_NUM
number_formats[34] = FMT_NUM
number_formats[72] = FMT_PCT
number_formats[73] = FMT_PCT
number_formats[74] = FMT_PCT
number_formats[75] = FMT_NUM
number_formats[76] = FMT_PCT
number_formats[77] = FMT_NUM
number_formats[78] = FMT_NUM
number_formats[79] = FMT_PCT
number_formats[80] = FMT_PCT
for i in range(12):
    number_formats[84 + i] = FMT_DEC

for row, fmt in number_formats.items():
    ws1.cell(row=row, column=2).number_format = fmt

for r in (5, 32, 49, 70, 82):
    ws1.merge_cells(f"A{r}:D{r}")
ws1.merge_cells("C83:D83")
for r in range(84, 96):
    ws1.merge_cells(f"C{r}:D{r}")

print("1_HYPOTHESES reconstruite.")

# =====================================================================
# 3) RECONSTRUCTION DE 2_PREVISION_CA
# =====================================================================
# Nouvelle disposition :
#   6-17   : synthese produits x12 (A..H)
#   18     : TOTAL synthese
#   20     : titre ventilation mensuelle
#   21     : header mois
#   22-33  : ventilation VOLUME produits x12 (A..N)
#   34     : TOTAL VOLUMES
#   35-46  : ventilation CA produits x12 (A..N)
#   47     : TOTAL CA MENSUEL
#   49     : label "CA Global"
#   50-52  : Annee1/2/3

for rng in list(ws2.merged_cells.ranges):
    ws2.unmerge_cells(str(rng))
for row in ws2.iter_rows(min_row=1, max_row=ws2.max_row, max_col=14):
    for cell in row:
        cell.value = None

ws2["A1"] = title2
ws2["A2"] = subtitle2
ws2["A4"] = titre_synth
set_row(ws2, 5, ["Gamme de Produit", "Prix Unitaire", "Volume Annee 1", "CA Annee 1 (FCFA)",
                  "Volume Annee 2", "CA Annee 2 (FCFA)", "Volume Annee 3", "CA Annee 3 (FCFA)"],
        style2_header)

col_L = get_column_letter
for i in range(N_PRODUITS):
    r = 6 + i
    vent_r = 22 + i
    hyp_prix_r = 6 + i
    name = synth_rows[i][0] if i < N_PRODUITS_ACTUELS else ""
    ws2.cell(row=r, column=1, value=name)
    apply_style(ws2.cell(row=r, column=1), style2_prod_label)
    formulas = [
        f"=1_HYPOTHESES!B{hyp_prix_r}",
        f"=N{vent_r}",
        f"=C{r}*B{r}",
        f"=C{r}*(1+1_HYPOTHESES!$B$30)",
        f"=E{r}*B{r}",
        f"=E{r}*(1+1_HYPOTHESES!$B$31)",
        f"=G{r}*B{r}",
    ]
    for j, f in enumerate(formulas, start=2):
        cell = ws2.cell(row=r, column=j, value=f)
        apply_style(cell, style2_num)
    # Colonnes Volume (C,E,G) : format nombre simple, pas FCFA
    for volcol in (3, 5, 7):
        ws2.cell(row=r, column=volcol).number_format = '#,##0'

ws2.cell(row=18, column=1, value="TOTAL CHIFFRE D'AFFAIRES")
apply_style(ws2.cell(row=18, column=1), style2_total[0])
for j, colletter in enumerate(["C", "D", "E", "F", "G", "H"], start=3):
    cell = ws2.cell(row=18, column=j, value=f"=SUM({colletter}6:{colletter}17)")
    apply_style(cell, style2_total[j - 1])

ws2["A20"] = titre_vent
set_row(ws2, 21, ["Produit / Indicateur"] + [f"M{m}" for m in range(1, 13)] + ["TOTAL A1"],
        [style2_month_header] * 14)

months_cols = [col_L(c) for c in range(2, 14)]  # B..M
for i in range(N_PRODUITS):
    r = 22 + i
    hyp_vol_r = 18 + i
    name = f"Vol. {synth_rows[i][0]} (unites)" if i < N_PRODUITS_ACTUELS else ""
    ws2.cell(row=r, column=1, value=name)
    apply_style(ws2.cell(row=r, column=1), style2_prod_label)
    for m, colletter in enumerate(months_cols, start=1):
        season_r = 83 + m
        f = f"=1_HYPOTHESES!$B${hyp_vol_r}*1_HYPOTHESES!$B${season_r}"
        cell = ws2[f"{colletter}{r}"] = f
        apply_style(ws2[f"{colletter}{r}"], style2_vent_num)
    cell = ws2.cell(row=r, column=14, value=f"=SUM(B{r}:M{r})")
    apply_style(cell, style2_vent_total)

ws2.cell(row=34, column=1, value="TOTAL VOLUMES VENDUS (unites)")
apply_style(ws2.cell(row=34, column=1), style2_vol_total_row)
for colletter in months_cols + ["N"]:
    cell = ws2[f"{colletter}34"] = f"=SUM({colletter}22:{colletter}33)"
    apply_style(ws2[f"{colletter}34"], style2_vol_total_row)
    ws2[f"{colletter}34"].number_format = '#,##0'

for i in range(N_PRODUITS):
    r = 35 + i
    vent_r = 22 + i
    hyp_prix_r = 6 + i
    name = f"CA {synth_rows[i][0]} (FCFA)" if i < N_PRODUITS_ACTUELS else ""
    ws2.cell(row=r, column=1, value=name)
    apply_style(ws2.cell(row=r, column=1), style2_prod_label)
    for colletter in months_cols:
        f = f"={colletter}{vent_r}*1_HYPOTHESES!$B${hyp_prix_r}"
        ws2[f"{colletter}{r}"] = f
        apply_style(ws2[f"{colletter}{r}"], style2_vent_num)
        ws2[f"{colletter}{r}"].number_format = '#,##0" FCFA"'
    cell = ws2.cell(row=r, column=14, value=f"=SUM(B{r}:M{r})")
    apply_style(cell, style2_vent_total)
    cell.number_format = '#,##0" FCFA"'

ws2.cell(row=47, column=1, value="TOTAL CHIFFRE D'AFFAIRES MENSUEL (FCFA)")
apply_style(ws2.cell(row=47, column=1), style2_ca_mensuel_total)
for colletter in months_cols:
    cell = ws2[f"{colletter}47"] = f"=SUM({colletter}35:{colletter}46)"
    apply_style(ws2[f"{colletter}47"], style2_ca_mensuel_total)
    ws2[f"{colletter}47"].number_format = '#,##0" FCFA"'
cell = ws2.cell(row=47, column=14, value="=SUM(B47:M47)")
apply_style(cell, style2_ca_mensuel_total)
cell.number_format = '#,##0" FCFA"'

ws2["B49"] = "CA Global"
ws2["A50"] = "Annee 1"
ws2["B50"] = "=D18"
ws2["A51"] = "Annee 2"
ws2["B51"] = "=F18"
ws2["A52"] = "Annee 3"
ws2["B52"] = "=H18"

# Le graphique "Evolution du Chiffre d'Affaires sur 3 Ans" pointait encore vers
# A22:A24/B22:B24 (position du bloc CA Global AVANT restructuration). Ce bloc a
# ete deplace en 50-52 ci-dessus mais la reference du graphique n'avait jamais
# ete mise a jour - exactement la classe de bug (reference codee en dur qui ne
# suit pas un decalage de lignes) que ce script doit precisement eviter.
for chart in ws2._charts:
    for series in chart.series:
        if series.val and series.val.numRef:
            series.val.numRef.f = "2_PREVISION_CA!$B$50:$B$52"
        if series.cat and series.cat.strRef:
            series.cat.strRef.f = "2_PREVISION_CA!$A$50:$A$52"

print("2_PREVISION_CA reconstruite.")

# =====================================================================
# 4) RECONSTRUCTION DE 4_INVESTISSEMENTS
# =====================================================================
# Nouvelle disposition :
#   6-20   : investissements x15
#   21     : TOTAL
#   23     : titre financement
#   24     : header financement
#   25-26  : Fonds Propres / Emprunt
#   27     : TOTAL RESSOURCES

for row in ws4.iter_rows(min_row=1, max_row=ws4.max_row, max_col=6):
    for cell in row:
        cell.value = None

ws4["A1"] = title4
ws4["A2"] = subtitle4
ws4["A4"] = titre_capex
set_row(ws4, 5, ["Poste d'Investissement", "Montant HT (FCFA)", "Duree (Ans)",
                  "Amort. Annuel (FCFA)", "Amort. Mensuel (FCFA)", "% du CapEx Total"],
        style4_header)

for i in range(N_INVEST):
    r = 6 + i
    if i < N_INVEST_ACTUELS:
        label, montant, duree = invest_rows[i][0], invest_rows[i][1], invest_rows[i][2]
    else:
        label, montant, duree = "", None, None
    ws4.cell(row=r, column=1, value=label)
    apply_style(ws4.cell(row=r, column=1), style4_label)
    ws4.cell(row=r, column=2, value=montant)
    apply_style(ws4.cell(row=r, column=2), style4_num)
    ws4.cell(row=r, column=3, value=duree)
    apply_style(ws4.cell(row=r, column=3), style4_num)
    ws4.cell(row=r, column=3).number_format = '#,##0'
    for j, f in ((4, f'=IF(C{r}=0,0,B{r}/C{r})'), (5, f"=D{r}/12"), (6, f"=B{r}/$B$21")):
        cell = ws4.cell(row=r, column=j, value=f)
        apply_style(cell, style4_num)
    ws4.cell(row=r, column=6).number_format = '0.0%'

ws4.cell(row=21, column=1, value="TOTAL INVESTISSEMENTS (CAPEX)")
apply_style(ws4.cell(row=21, column=1), style4_total[0])
for j, colletter in ((2, "B"), (4, "D"), (5, "E"), (6, "F")):
    cell = ws4.cell(row=21, column=j, value=f"=SUM({colletter}6:{colletter}20)")
    apply_style(cell, style4_total[j - 1])
ws4.cell(row=21, column=6).number_format = '0.0%'

ws4["A23"] = titre_financement
set_row(ws4, 24, ["Source de Financement", "% Financement", "Montant (FCFA)", "Conditions / Modalites"],
        style4_fin_header)

set_row(ws4, 25, [row17[0], "=1_HYPOTHESES!B72", "=B21*B25", row17[3]], style4_fin_line)
set_row(ws4, 26, [row18[0], "=1_HYPOTHESES!B73", "=B21*B26",
                   "Taux '1_HYPOTHESES'!B74 sur '1_HYPOTHESES'!B75 mois"], style4_fin_line)
set_row(ws4, 27, ["TOTAL RESSOURCES DE FINANCEMENT", "=SUM(B25:B26)", "=SUM(C25:C26)", ""], style4_fin_total)

print("4_INVESTISSEMENTS reconstruite.")

# =====================================================================
# 5) MISE A JOUR DE 3_COMPTE_RESULTAT
# =====================================================================
# Lignes 1-10 inchangees (structure et formules internes identiques).
# Lignes 11-17 (7 categories OPEX nommees) fusionnees en UNE ligne
# generique 11 (decision validee), le reste (EBITDA..Resultat net) remonte
# de 6 lignes : 19->12, 20->13, 21->14, 22->15, 23->16, 24->17, 25->18.

ws3 = wb["3_COMPTE_RESULTAT"]

style3_total_row = [style_of(ws3, c) for c in ("A10", "B10", "C10", "D10", "E10", "F10", "G10")]
style3_normal_row = [style_of(ws3, c) for c in ("A11", "B11", "C11", "D11", "E11", "F11", "G11")]

# Sauvegarde des libelles avant ecrasement (rows 19-25 deviennent 12-18)
labels_bas = {old: ws3.cell(row=old, column=1).value for old in range(19, 26)}

# Effacer les lignes 11 a 25 (bloc a reconstruire), styles geres explicitement
for r in range(11, 26):
    for c in range(1, 8):
        ws3.cell(row=r, column=c).value = None

# Ligne 11 : TOTAL CHARGES FIXES (OPEX) generique
ws3.cell(row=11, column=1, value="TOTAL CHARGES FIXES (OPEX)")
formulas_11 = [
    None,
    "=SUM(1_HYPOTHESES!B50:B69)*12",
    "=B11/$B$5",
    "=B11*(1+1_HYPOTHESES!$B$79)",
    "=D11/$D$5",
    "=B11*(1+1_HYPOTHESES!$B$80)",
    "=F11/$F$5",
]
for j, f in enumerate(formulas_11, start=1):
    if f is None:
        continue
    ws3.cell(row=11, column=j, value=f)
for j in range(1, 8):
    apply_style(ws3.cell(row=11, column=j), style3_total_row[j - 1])

# Ligne 12 : EBITDA (etait 19)
ws3.cell(row=12, column=1, value=labels_bas[19])
for j, f in zip(range(2, 8), ["=B10-B11", "=B12/$B$5", "=D10-D11", "=D12/$D$5", "=F10-F11", "=F12/$F$5"]):
    ws3.cell(row=12, column=j, value=f)
for j in range(1, 8):
    apply_style(ws3.cell(row=12, column=j), style3_total_row[j - 1])

# Ligne 13 : Dotations aux Amortissements (etait 20) - reference 4_INVESTISSEMENTS!D21 (etait D13)
ws3.cell(row=13, column=1, value=labels_bas[20])
for j, f in zip(range(2, 8), ["=4_INVESTISSEMENTS!D21", "=B13/$B$5", "=4_INVESTISSEMENTS!D21",
                                "=D13/$D$5", "=4_INVESTISSEMENTS!D21", "=F13/$F$5"]):
    ws3.cell(row=13, column=j, value=f)
for j in range(1, 8):
    apply_style(ws3.cell(row=13, column=j), style3_normal_row[j - 1])

# Ligne 14 : RESULTAT OPERATIONNEL (EBIT) (etait 21)
ws3.cell(row=14, column=1, value=labels_bas[21])
for j, f in zip(range(2, 8), ["=B12-B13", "=B14/$B$5", "=D12-D13", "=D14/$D$5", "=F12-F13", "=F14/$F$5"]):
    ws3.cell(row=14, column=j, value=f)
for j in range(1, 8):
    apply_style(ws3.cell(row=14, column=j), style3_total_row[j - 1])

# Ligne 15 : Charges Financieres (etait 22) - CUMIPMT reference 1_HYPOTHESES!B74/B75 et 4_INVESTISSEMENTS!C26 (etaient B35/B36, C18)
ws3.cell(row=15, column=1, value=labels_bas[22])
formulas_15 = {
    2: "=-CUMIPMT(1_HYPOTHESES!$B$74/12,1_HYPOTHESES!$B$75,4_INVESTISSEMENTS!$C$26,1,12,0)",
    4: "=-CUMIPMT(1_HYPOTHESES!$B$74/12,1_HYPOTHESES!$B$75,4_INVESTISSEMENTS!$C$26,13,24,0)",
    6: "=-CUMIPMT(1_HYPOTHESES!$B$74/12,1_HYPOTHESES!$B$75,4_INVESTISSEMENTS!$C$26,25,36,0)",
    3: "=B15/$B$5",
    5: "=D15/$D$5",
    7: "=F15/$F$5",
}
for j, f in formulas_15.items():
    ws3.cell(row=15, column=j, value=f)
for j in range(1, 8):
    apply_style(ws3.cell(row=15, column=j), style3_normal_row[j - 1])

# Ligne 16 : RESULTAT AVANT IMPOT (EBT) (etait 23)
ws3.cell(row=16, column=1, value=labels_bas[23])
for j, f in zip(range(2, 8), ["=B14-B15", "=B16/$B$5", "=D14-D15", "=D16/$D$5", "=F14-F15", "=F16/$F$5"]):
    ws3.cell(row=16, column=j, value=f)
for j in range(1, 8):
    apply_style(ws3.cell(row=16, column=j), style3_normal_row[j - 1])

# Ligne 17 : Impot sur les Societes (IS) (etait 24) - reference 1_HYPOTHESES!B76 (etait B37)
ws3.cell(row=17, column=1, value=labels_bas[24])
formulas_17 = {
    2: "=IF(B16>0, B16*1_HYPOTHESES!$B$76, 0)",
    4: "=IF(D16>0, D16*1_HYPOTHESES!$B$76, 0)",
    6: "=IF(F16>0, F16*1_HYPOTHESES!$B$76, 0)",
    3: "=B17/$B$5",
    5: "=D17/$D$5",
    7: "=F17/$F$5",
}
for j, f in formulas_17.items():
    ws3.cell(row=17, column=j, value=f)
for j in range(1, 8):
    apply_style(ws3.cell(row=17, column=j), style3_normal_row[j - 1])

# Ligne 18 : RESULTAT NET COMPTABLE (etait 25)
ws3.cell(row=18, column=1, value=labels_bas[25])
for j, f in zip(range(2, 8), ["=B16-B17", "=B18/$B$5", "=D16-D17", "=D18/$D$5", "=F16-F17", "=F18/$F$5"]):
    ws3.cell(row=18, column=j, value=f)
for j in range(1, 8):
    apply_style(ws3.cell(row=18, column=j), style3_total_row[j - 1])

# Mise a jour des formules amont (lignes 6,7,8 - couts variables) pour generaliser
# sur la plage de 12 produits au lieu de 3 references nommees
ws3["B6"] = "=SUMPRODUCT(2_PREVISION_CA!C6:C17,1_HYPOTHESES!B35:B46)"
ws3["D6"] = "=SUMPRODUCT(2_PREVISION_CA!E6:E17,1_HYPOTHESES!B35:B46)"
ws3["F6"] = "=SUMPRODUCT(2_PREVISION_CA!G6:G17,1_HYPOTHESES!B35:B46)"
ws3["B7"] = "=2_PREVISION_CA!C18*1_HYPOTHESES!B47"
ws3["D7"] = "=2_PREVISION_CA!E18*1_HYPOTHESES!B47"
ws3["F7"] = "=2_PREVISION_CA!G18*1_HYPOTHESES!B47"
ws3["B8"] = "=2_PREVISION_CA!C18*1_HYPOTHESES!B48"
ws3["D8"] = "=2_PREVISION_CA!E18*1_HYPOTHESES!B48"
ws3["F8"] = "=2_PREVISION_CA!G18*1_HYPOTHESES!B48"

# Lignes 5,9,10 : formules internes (2_PREVISION_CA!D9/F9/H9 -> D18/F18/H18)
ws3["B5"] = "=2_PREVISION_CA!D18"
ws3["D5"] = "=2_PREVISION_CA!F18"
ws3["F5"] = "=2_PREVISION_CA!H18"

print("3_COMPTE_RESULTAT mise a jour.")

# =====================================================================
# 6) MISE A JOUR DE 5_TRESORERIE
# =====================================================================
ws5 = wb["5_TRESORERIE"]

# Ligne 6 : encaissements ventes clients (delai 30j) - reference 2_PREVISION_CA!col47 (etait col20)
ws5["B6"] = "=2_PREVISION_CA!B47*0.5"
for i, colletter in enumerate(months_cols[1:], start=1):
    prev = months_cols[i - 1]
    ws5[f"{colletter}6"] = f"=(2_PREVISION_CA!{prev}47*0.5)+(2_PREVISION_CA!{colletter}47*0.5)"

# Ligne 7 : apport en capital (etait 4_INVESTISSEMENTS!C17 -> C25)
ws5["B7"] = "=4_INVESTISSEMENTS!C25"
# Ligne 8 : deblocage emprunt (etait 4_INVESTISSEMENTS!C18 -> C26)
ws5["B8"] = "=4_INVESTISSEMENTS!C26"

# Ligne 10 : decaissements achats MP & packaging - generalise sur 12 produits
for colletter in months_cols:
    ws5[f"{colletter}10"] = (
        f"=(SUMPRODUCT(2_PREVISION_CA!{colletter}22:{colletter}33,1_HYPOTHESES!$B$35:$B$46)"
        f"+2_PREVISION_CA!{colletter}34*1_HYPOTHESES!$B$47)*0.7"
    )

# Ligne 11 : decaissements OPEX fixes (mensuel, deja monthly, pas de *12) - generalise sur 20 charges
for colletter in months_cols:
    ws5[f"{colletter}11"] = "=SUM(1_HYPOTHESES!$B$50:$B$69)"
for colletter in [col_L(c) for c in range(14, 26)]:  # N..Y (mois 13-24, annee 2)
    ws5[f"{colletter}11"] = "=(3_COMPTE_RESULTAT!$D$11/12)"
for colletter in [col_L(c) for c in range(26, 38)]:  # Z..AK (mois 25-36, annee 3)
    ws5[f"{colletter}11"] = "=(3_COMPTE_RESULTAT!$F$11/12)"

# Ligne 12 : investissement initial CapEx (etait 4_INVESTISSEMENTS!$B$13 -> $B$21)
ws5["B12"] = "=4_INVESTISSEMENTS!$B$21"

# Ligne 13 : service de la dette (PMT) - taux/duree deplaces, emprunt C18->C26
for colletter in [col_L(c) for c in range(2, 38)]:  # B..AK
    ws5[f"{colletter}13"] = "=PMT(1_HYPOTHESES!$B$74/12, 1_HYPOTHESES!$B$75, -4_INVESTISSEMENTS!$C$26)"

# Ligne 14 : paiement IS (etait 3_COMPTE_RESULTAT!$B$24/$D$24 -> $B$17/$D$17)
ws5["P14"] = "=3_COMPTE_RESULTAT!$B$17"
ws5["AB14"] = "=3_COMPTE_RESULTAT!$D$17"

print("5_TRESORERIE mise a jour.")

# =====================================================================
# 7) MISE A JOUR DE 7_RENTABILITE
# =====================================================================
ws7 = wb["7_RENTABILITE"]
ws7["B10"] = "=3_COMPTE_RESULTAT!B11+3_COMPTE_RESULTAT!B13"
ws7["B12"] = "=3_COMPTE_RESULTAT!B5/2_PREVISION_CA!C18"

print("7_RENTABILITE mise a jour.")

# =====================================================================
# 6bis) MISE A JOUR DE 6_BFR
# =====================================================================
# Delai clients (etait 1_HYPOTHESES!B38 -> B77), delai fournisseurs (etait B39 -> B78)
ws6 = wb["6_BFR"]
ws6["B7"] = "=3_COMPTE_RESULTAT!B5*(1_HYPOTHESES!$B$77/365)"
ws6["C7"] = "=3_COMPTE_RESULTAT!D5*(1_HYPOTHESES!$B$77/365)"
ws6["D7"] = "=3_COMPTE_RESULTAT!F5*(1_HYPOTHESES!$B$77/365)"
ws6["B9"] = "=(3_COMPTE_RESULTAT!B6+3_COMPTE_RESULTAT!B7)*(1_HYPOTHESES!$B$78/365)"
ws6["C9"] = "=(3_COMPTE_RESULTAT!D6+3_COMPTE_RESULTAT!D7)*(1_HYPOTHESES!$B$78/365)"
ws6["D9"] = "=(3_COMPTE_RESULTAT!F6+3_COMPTE_RESULTAT!F7)*(1_HYPOTHESES!$B$78/365)"

print("6_BFR mise a jour.")

# =====================================================================
# 8) MISE A JOUR DE 8_SCENARIOS
# =====================================================================
ws8 = wb["8_SCENARIOS"]
ws8["B7"] = "=3_COMPTE_RESULTAT!D11"
ws8["C7"] = "=3_COMPTE_RESULTAT!D11"
ws8["D7"] = "=3_COMPTE_RESULTAT!D11*1.05"
ws8["B10"] = "=(B9-3_COMPTE_RESULTAT!D13-3_COMPTE_RESULTAT!D15)*(1-1_HYPOTHESES!$B$76)"
ws8["C10"] = "=3_COMPTE_RESULTAT!D18"
ws8["D10"] = "=(D9-3_COMPTE_RESULTAT!D13-3_COMPTE_RESULTAT!D15)*(1-1_HYPOTHESES!$B$76)"

print("8_SCENARIOS mise a jour.")

# =====================================================================
# 9) MISE A JOUR DE 9_DASHBOARD
# =====================================================================
ws9 = wb["9_DASHBOARD"]
ws9["B5"] = "=2_PREVISION_CA!D18"
ws9["E5"] = "=(2_PREVISION_CA!F18-2_PREVISION_CA!D18)/2_PREVISION_CA!D18"
ws9["B8"] = "=3_COMPTE_RESULTAT!B12"
ws9["E8"] = "=3_COMPTE_RESULTAT!B18"
ws9["B18"] = "=3_COMPTE_RESULTAT!B12"
ws9["C18"] = "=3_COMPTE_RESULTAT!D12"
ws9["D18"] = "=3_COMPTE_RESULTAT!F12"
ws9["B19"] = "=3_COMPTE_RESULTAT!B18"
ws9["C19"] = "=3_COMPTE_RESULTAT!D18"
ws9["D19"] = "=3_COMPTE_RESULTAT!F18"

print("9_DASHBOARD mise a jour.")

wb.save(DST)
print("Classeur restructure sauvegarde:", DST)
