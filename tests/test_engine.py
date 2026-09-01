"""
Suite de tests du moteur de generation (Lot 1).

Verifie, pour 3 profils de projet differents :
  - generation + recalcul sans erreur
  - 11 feuilles, dans le meme ordre que le gabarit de reference
  - zero erreur de formule (#REF!, #DIV/0!, ...) apres recalcul
  - coherence de signe / plausibilite des totaux (CA, EBITDA, Resultat net,
    tresorerie croissants ou du moins non aberrants)
  - aucune formule du gabarit n'a ete modifiee par le moteur (comparaison
    des cellules-formules avant/apres ecriture des donnees du projet)
"""
import json
import os

import openpyxl
import pytest

from engine.generator import generate_workbook, TEMPLATE_PATH
from engine.models import ProjetInput
from engine.recalc import recalculate_and_check, scan_formula_errors, RecalcError
from engine.structure_check import EXPECTED_SHEETS, check_sheet_order

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, "data")
PROFILES = [
    "exemple_mono_produit.json",
    "exemple_multi_produits_saisonnalite.json",
    "exemple_investissement_lourd.json",
]


def _load_projet(filename: str) -> ProjetInput:
    with open(os.path.join(DATA_DIR, filename), encoding="utf-8") as f:
        data = json.load(f)
    return ProjetInput.from_dict(data)


def _snapshot_template_formulas() -> dict:
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    snap = {}
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and v.startswith("="):
                    snap[(sheet, cell.coordinate)] = v
    return snap


@pytest.mark.parametrize("profile", PROFILES)
def test_generation_recalcul_sans_erreur(profile, tmp_path):
    projet = _load_projet(profile)
    out_path = str(tmp_path / f"{profile}.xlsx")

    generate_workbook(projet, out_path)
    recalculate_and_check(out_path)  # leve RecalcError si erreurs -> le test echoue

    errors = scan_formula_errors(out_path)
    assert errors == [], f"Erreurs de formule detectees pour {profile}: {errors}"


@pytest.mark.parametrize("profile", PROFILES)
def test_structure_11_feuilles_meme_ordre(profile, tmp_path):
    projet = _load_projet(profile)
    out_path = str(tmp_path / f"{profile}.xlsx")
    generate_workbook(projet, out_path)

    check_sheet_order(out_path)  # leve AssertionError si incorrect
    wb = openpyxl.load_workbook(out_path, read_only=True)
    assert len(wb.sheetnames) == 11
    assert wb.sheetnames == EXPECTED_SHEETS


@pytest.mark.parametrize("profile", PROFILES)
def test_aucune_formule_modifiee(profile, tmp_path):
    before = _snapshot_template_formulas()

    projet = _load_projet(profile)
    out_path = str(tmp_path / f"{profile}.xlsx")
    generate_workbook(projet, out_path)  # leve RuntimeError en interne si violation

    wb = openpyxl.load_workbook(out_path)
    after = {}
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and v.startswith("="):
                    after[(sheet, cell.coordinate)] = v

    for key, formula in before.items():
        assert after.get(key) == formula, f"Formule modifiee en {key}: {formula!r} -> {after.get(key)!r}"


@pytest.mark.parametrize("profile", PROFILES)
def test_coherence_financiere_plausible(profile, tmp_path):
    projet = _load_projet(profile)
    out_path = str(tmp_path / f"{profile}.xlsx")
    generate_workbook(projet, out_path)
    recalculate_and_check(out_path)

    wb = openpyxl.load_workbook(out_path, data_only=True)
    cr = wb["3_COMPTE_RESULTAT"]
    ca = [cr["B5"].value, cr["D5"].value, cr["F5"].value]
    resnet = [cr["B18"].value, cr["D18"].value, cr["F18"].value]
    tresorerie_fin = wb["5_TRESORERIE"]["AK16"].value

    assert all(v is not None for v in ca), "CA manquant sur une des 3 annees"
    assert all(v > 0 for v in ca), f"CA doit etre positif sur les 3 annees: {ca}"
    # Croissance : le CA ne doit jamais regresser d'une annee sur l'autre
    assert ca[1] >= ca[0] and ca[2] >= ca[1], f"CA doit croitre ou rester stable: {ca}"
    assert resnet[2] > resnet[0], f"Le resultat net doit s'ameliorer sur la duree du projet: {resnet}"
    assert tresorerie_fin is not None


def test_validation_rejette_projet_sans_produit():
    from engine.models import ProjetInput as PI, ProjetInputError
    data = {
        "secteur_activite": "Test",
        "produits": [],
        "charges_fixes": [{"nom": "Loyer", "montant_mensuel": 100000}],
        "taux_croissance_a2": 0.1,
        "taux_croissance_a3": 0.1,
    }
    projet = PI.from_dict(data)
    with pytest.raises(ProjetInputError):
        projet.validate()


def test_produit_cout_variable_calcule_depuis_sous_postes():
    from engine.models import Produit

    p = Produit(
        nom="Test",
        prix_unitaire=100,
        volume_initial_mensuel=10,
        cout_matiere_premiere=30,
        cout_emballage=15,
        cout_production=5,
    )
    assert p.cout_variable_unitaire == 50


def test_metadonnees_non_ecrites_dans_le_classeur(tmp_path):
    projet = _load_projet("exemple_mono_produit.json")
    projet.nom_projet = "Ma Super PME"
    projet.forme_juridique = "SARL"
    out_path = str(tmp_path / "meta.xlsx")
    generate_workbook(projet, out_path)

    wb = openpyxl.load_workbook(out_path)
    found = []
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        for row in ws.iter_rows():
            for cell in row:
                if cell.value == "Ma Super PME" or cell.value == "SARL":
                    found.append((sheet, cell.coordinate, cell.value))
    assert found == [], f"Des metadonnees ont ete ecrites dans le classeur: {found}"


def test_graphique_ca_reference_le_bon_bloc_ca_global():
    """Regression : le graphique 'Evolution du CA sur 3 Ans' de 2_PREVISION_CA
    doit pointer sur le bloc CA Global (lignes 50-52), pas sur un bloc deplace
    lors d'une restructuration anterieure du gabarit (cf. cahier des charges :
    reference codee en dur qui ne suit pas un decalage de lignes)."""
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws = wb["2_PREVISION_CA"]
    assert ws["A50"].value == "Annee 1" and ws["A52"].value == "Annee 3"

    charts = ws._charts
    assert len(charts) == 1, "Le gabarit doit conserver exactement 1 graphique sur 2_PREVISION_CA"
    series = charts[0].series[0]
    assert series.val.numRef.f == "2_PREVISION_CA!$B$50:$B$52"
    assert series.cat.strRef.f == "2_PREVISION_CA!$A$50:$A$52"


def test_validation_rejette_trop_de_produits():
    from engine.models import ProjetInput as PI, ProjetInputError
    data = {
        "secteur_activite": "Test",
        "produits": [
            {"nom": f"P{i}", "prix_unitaire": 100, "volume_initial_mensuel": 10, "cout_variable_unitaire": 50}
            for i in range(13)
        ],
        "charges_fixes": [{"nom": "Loyer", "montant_mensuel": 100000}],
        "taux_croissance_a2": 0.1,
        "taux_croissance_a3": 0.1,
    }
    projet = PI.from_dict(data)
    with pytest.raises(ProjetInputError):
        projet.validate()
