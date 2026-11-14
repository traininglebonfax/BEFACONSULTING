"""
Tests de engine/preview.py (Lot 3 - apercu par feuille).

Verifie que la conversion HTML lit bien les valeurs recalculees (jamais une
formule brute) et que le format FCFA / la structure 11 feuilles sont
preserves.
"""
import json
import os

import pytest

from engine.generator import generate_workbook
from engine.models import ProjetInput
from engine.preview import workbook_to_preview
from engine.recalc import recalculate_and_check
from engine.structure_check import EXPECTED_SHEETS

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, "data")


def _load_projet(filename: str) -> ProjetInput:
    with open(os.path.join(DATA_DIR, filename), encoding="utf-8") as f:
        data = json.load(f)
    return ProjetInput.from_dict(data)


@pytest.fixture(scope="module")
def recalculated_preview(tmp_path_factory):
    projet = _load_projet("exemple_mono_produit.json")
    out_path = str(tmp_path_factory.mktemp("preview") / "preview.xlsx")
    generate_workbook(projet, out_path)
    recalculate_and_check(out_path)
    return workbook_to_preview(out_path)


def test_11_feuilles_dans_le_bon_ordre(recalculated_preview):
    names = [s["name"] for s in recalculated_preview]
    assert names == EXPECTED_SHEETS


def test_html_ne_contient_aucune_formule_brute(recalculated_preview):
    for sheet in recalculated_preview:
        # Une formule brute non evaluee ressemblerait a "=SOMME(" ou "=B5*"
        assert "<td>=" not in sheet["html"], f"Formule brute trouvee dans {sheet['name']}"


def test_html_contient_les_valeurs_recalculees(recalculated_preview):
    hypotheses = next(s for s in recalculated_preview if s["name"] == "1_HYPOTHESES")
    assert "<table" in hypotheses["html"]
    assert "<td" in hypotheses["html"]
    # Le nom du produit du jeu de donnees mono-produit doit apparaitre dans l'apercu
    with open(os.path.join(DATA_DIR, "exemple_mono_produit.json"), encoding="utf-8") as f:
        data = json.load(f)
    nom_produit = data["produits"][0]["nom"]
    assert nom_produit in hypotheses["html"]


def test_format_fcfa_preserve(recalculated_preview):
    compte_resultat = next(s for s in recalculated_preview if s["name"] == "3_COMPTE_RESULTAT")
    assert "FCFA" in compte_resultat["html"]


def test_graphique_ca_rendu_en_svg(recalculated_preview):
    ca_sheet = next(s for s in recalculated_preview if s["name"] == "2_PREVISION_CA")
    assert 'class="bp-chart"' in ca_sheet["html"]
    assert ca_sheet["html"].count("<rect") == 3  # 3 barres : Annee 1/2/3


def test_ai_suggested_cells_ajoute_un_badge_dans_la_bonne_feuille():
    projet = _load_projet("exemple_mono_produit.json")
    from tempfile import mkdtemp
    out_path = os.path.join(mkdtemp(), "preview_sugg.xlsx")
    generate_workbook(projet, out_path)
    recalculate_and_check(out_path)

    sheets = workbook_to_preview(
        out_path, {("1_HYPOTHESES", "B35"): "Cout variable suggere par l'IA : 450 FCFA"}
    )
    hypotheses = next(s for s in sheets if s["name"] == "1_HYPOTHESES")
    autre = next(s for s in sheets if s["name"] == "2_PREVISION_CA")
    assert "bp-ai-suggested" in hypotheses["html"]
    assert "✦ IA" in hypotheses["html"]
    assert "bp-ai-suggested" not in autre["html"]


def test_statuts_dashboard_colores_sans_emoji_brut(recalculated_preview):
    """Regression : les pastilles de statut (Statut/Tendance de 9_DASHBOARD)
    ne s'affichaient pas en couleur dans le navigateur (rendu emoji non fiable
    selon la police du systeme) - elles doivent etre remplacees par un glyphe
    texte simple dont la couleur est fixee en CSS."""
    dashboard = next(s for s in recalculated_preview if s["name"] == "9_DASHBOARD")
    assert 'style="color:#1E4D2B;font-weight:700">●</span>' in dashboard["html"]
    for emoji in ("\U0001F7E2", "\U0001F7E1", "\U0001F534"):
        assert emoji not in dashboard["html"]
