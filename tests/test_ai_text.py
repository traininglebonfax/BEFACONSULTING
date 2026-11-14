"""
Tests de engine/ai_text.py et engine/generator.attach_ai_comments (Lot 4).

Aucun de ces tests n'appelle reellement l'API Anthropic (pas de cle API
requise en CI) : on verifie la gestion d'absence de cle, le parsing de
reponse, et l'attache des commentaires dans le classeur.
"""
import json
import os

import openpyxl
import pytest

from engine.ai_text import (
    AiTextError,
    _parse_cost_suggestions,
    _parse_response,
    generate_narrative_texts,
    suggest_variable_costs,
)
from engine.generator import attach_ai_comments, attach_cost_suggestion_comments, generate_workbook
from engine.models import ProjetInput

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, "data")


def _load_projet(filename: str) -> ProjetInput:
    with open(os.path.join(DATA_DIR, filename), encoding="utf-8") as f:
        data = json.load(f)
    return ProjetInput.from_dict(data)


def test_generate_narrative_texts_leve_erreur_sans_cle_api(monkeypatch):
    monkeypatch.delenv("ANTHROPIC_API_KEY", raising=False)
    monkeypatch.delenv("ANTHROPIC_AUTH_TOKEN", raising=False)
    projet = _load_projet("exemple_mono_produit.json")
    with pytest.raises(AiTextError):
        generate_narrative_texts(projet)


def test_parse_response_extrait_les_deux_textes():
    text = (
        "DESCRIPTION_ACTIVITE: Une entreprise de jus naturels en Cote d'Ivoire.\n"
        "JUSTIFICATION_CROISSANCE: La croissance est portee par la demande locale."
    )
    result = _parse_response(text)
    assert result["description_activite"] == "Une entreprise de jus naturels en Cote d'Ivoire."
    assert result["justification_croissance"] == "La croissance est portee par la demande locale."


def test_parse_response_gere_texte_multiligne():
    text = (
        "DESCRIPTION_ACTIVITE: Premiere phrase.\n"
        "Deuxieme phrase sur une nouvelle ligne.\n"
        "JUSTIFICATION_CROISSANCE: Justification en une ligne."
    )
    result = _parse_response(text)
    assert result["description_activite"] == "Premiere phrase. Deuxieme phrase sur une nouvelle ligne."
    assert result["justification_croissance"] == "Justification en une ligne."


def test_attach_ai_comments_ecrit_les_commentaires_sans_toucher_aux_valeurs(tmp_path):
    projet = _load_projet("exemple_mono_produit.json")
    out_path = str(tmp_path / "ai_comments.xlsx")
    generate_workbook(projet, out_path)

    wb_before = openpyxl.load_workbook(out_path)
    valeur_a1_avant = wb_before["1_HYPOTHESES"]["A1"].value
    valeur_a30_avant = wb_before["1_HYPOTHESES"]["A30"].value

    attach_ai_comments(out_path, "Description generee par test.", "Justification generee par test.")

    wb_after = openpyxl.load_workbook(out_path)
    ws = wb_after["1_HYPOTHESES"]
    assert ws["A1"].value == valeur_a1_avant
    assert ws["A30"].value == valeur_a30_avant
    assert ws["A1"].comment is not None
    assert "Description generee par test." in ws["A1"].comment.text
    assert ws["A30"].comment is not None
    assert "Justification generee par test." in ws["A30"].comment.text


def test_attach_ai_comments_ne_fait_rien_si_textes_vides(tmp_path):
    projet = _load_projet("exemple_mono_produit.json")
    out_path = str(tmp_path / "ai_comments_vide.xlsx")
    generate_workbook(projet, out_path)

    attach_ai_comments(out_path, "", "")

    wb = openpyxl.load_workbook(out_path)
    ws = wb["1_HYPOTHESES"]
    assert ws["A1"].comment is None
    assert ws["A30"].comment is None


# --- Suggestion de couts variables (role optionnel du Lot 4, cahier des charges Section2) ---

def test_suggest_variable_costs_leve_erreur_sans_cle_api(monkeypatch):
    monkeypatch.delenv("ANTHROPIC_API_KEY", raising=False)
    monkeypatch.delenv("ANTHROPIC_AUTH_TOKEN", raising=False)
    projet = _load_projet("exemple_mono_produit.json")
    with pytest.raises(AiTextError):
        suggest_variable_costs(projet, [0])


def test_suggest_variable_costs_retourne_vide_sans_index_manquant():
    projet = _load_projet("exemple_mono_produit.json")
    assert suggest_variable_costs(projet, []) == {}


def test_parse_cost_suggestions_extrait_les_valeurs():
    text = "0: 450\n1: 620 FCFA\n"
    result = _parse_cost_suggestions(text, [0, 1])
    assert result == {0: 450.0, 1: 620.0}


def test_parse_cost_suggestions_ignore_les_index_non_demandes():
    text = "0: 450\n5: 999\n"
    result = _parse_cost_suggestions(text, [0])
    assert result == {0: 450.0}


def test_attach_cost_suggestion_comments_ecrit_le_commentaire_sans_toucher_a_la_valeur(tmp_path):
    projet = _load_projet("exemple_mono_produit.json")
    out_path = str(tmp_path / "cost_comments.xlsx")
    generate_workbook(projet, out_path)

    wb_before = openpyxl.load_workbook(out_path)
    valeur_avant = wb_before["1_HYPOTHESES"]["B35"].value

    attach_cost_suggestion_comments(out_path, {0: 450.0})

    wb_after = openpyxl.load_workbook(out_path)
    cell = wb_after["1_HYPOTHESES"]["B35"]
    assert cell.value == valeur_avant
    assert cell.comment is not None
    assert "450" in cell.comment.text


def test_attach_cost_suggestion_comments_ne_fait_rien_si_vide(tmp_path):
    projet = _load_projet("exemple_mono_produit.json")
    out_path = str(tmp_path / "cost_comments_vide.xlsx")
    generate_workbook(projet, out_path)

    attach_cost_suggestion_comments(out_path, {})

    wb = openpyxl.load_workbook(out_path)
    assert wb["1_HYPOTHESES"]["B35"].comment is None
