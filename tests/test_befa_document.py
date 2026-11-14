"""
Lot 5 BEFA - document Word + PDF.

Generation reelle (LibreOffice requis) des deux exemples :
  - COREAL (nouveau projet) avec la redaction IA reelle enregistree ;
  - O'IVOIRE (entreprise existante) sans redaction : encarts "A completer".
"""
import json
import os
import re
import unicodedata

import pytest
from pypdf import PdfReader

from befa import document as D
from befa.dossier import DossierBEFA
from befa.projection import calculer_projection
from befa.redaction import Source, redaction_depuis_json
from engine.recalc import RecalcError, find_soffice

RACINE = os.path.dirname(os.path.dirname(__file__))

try:
    find_soffice()
    LIBREOFFICE = True
except RecalcError:
    LIBREOFFICE = False

avec_libreoffice = pytest.mark.skipif(not LIBREOFFICE, reason="LibreOffice absent")


def textes(pdf):
    return [unicodedata.normalize("NFKC", p.extract_text() or "") for p in PdfReader(pdf).pages]


def verifier_sommaire(pdf, sommaire):
    """Chaque titre du sommaire figure bien a la page indiquee."""
    pages = textes(pdf)
    sommaire_txt = " ".join(pages[1:3])
    for e in sommaire:
        m = re.search(re.escape(e.titre) + r"\s*\.*\s*(\d+)", sommaire_txt)
        assert m, f"titre absent du sommaire : {e.titre}"
        numero = int(m.group(1))
        assert D._normaliser(e.titre) in D._normaliser(pages[numero - 1]), f"{e.titre} absent de la page {numero}"


@avec_libreoffice
def test_business_plan_coreal(tmp_path):
    from scripts.generer_exemple_befa import exemple_coreal

    d, red, t = exemple_coreal()
    # un renvoi vers une source : appel de note [1] et bibliographie
    red.sections[0].paragraphes[0] += " [S1]"
    red.sources = [Source("S1", "USDA FAS — Oilseeds and Products Annual", "https://apps.fas.usda.gov/rapport")]
    res = D.generer(d, red, t, str(tmp_path), "bp")
    assert res.sommaire_complet and res.alertes == []
    assert 15 <= res.pages <= 30
    assert os.path.exists(res.docx)

    pages = textes(res.pdf)
    assert "COREAL FOOD INDUSTRY SA" in pages[0] and "BUSINESS PLAN" in pages[0]
    assert all(D.MARQUE_SOMMAIRE not in p for p in pages)
    assert all("Réalisé par Bureau d'Études Financières et d'Assistance-Consulting" in p for p in pages[1:])
    tout = "\n".join(pages)
    for attendu in ("COÛT TOTAL DES INVESTISSEMENTS", "Compte d'exploitation prévisionnel", "Plan de financement",
                    "Ratio de couverture de la dette", "Tendances de recherche en ligne (Google Trends)",
                    "Source : Google Trends, Côte d’Ivoire", "Analyse SWOT", "Bilan d'ouverture au 2024",
                    "[1]  USDA FAS — Oilseeds and Products Annual"):
        assert attendu in tout, attendu
    R = D.construire(d, red, calculer_projection(d.projection), t, {})
    verifier_sommaire(res.pdf, R.sommaire)


@avec_libreoffice
def test_dossier_credit_oivoire_sans_redaction(tmp_path):
    with open(os.path.join(RACINE, "data", "befa_exemple_oivoire.json"), encoding="utf-8") as f:
        d = DossierBEFA.from_dict(json.load(f))
    res = D.generer(d, None, None, str(tmp_path), "oivoire")
    assert res.sommaire_complet
    tout = "\n".join(textes(res.pdf))
    assert "DOSSIER DE CRÉDIT" in tout and "DEUXIÈME PARTIE : ÉTUDE FINANCIÈRE" in tout
    assert "[À compléter par le consultant" in tout  # sections narratives absentes
    assert "le fonds de roulement" in tout.lower() and "Analyse du bilan 2025" in tout
    assert "Augmentation du BFR" in tout


def test_formats_nombres():
    assert D._nombre(2_700_702_260) == "2 700 702 260"
    assert D._nombre(0) == "–" and D._nombre(None) == ""
    assert D._nombre(-1_144_423) == "−1 144 423"
    assert D._nombre(7.17, "ratio") == "7,2" and D._nombre(-0.28, "pct") == "−28 %"


def test_redaction_aller_retour_json():
    from scripts.generer_exemple_befa import exemple_coreal

    _, red, _ = exemple_coreal()
    copie = redaction_depuis_json(json.loads(json.dumps(red.vers_json())))
    assert copie.vers_json() == red.vers_json()


@avec_libreoffice
def test_routes_generer_et_telecharger():
    from fastapi.testclient import TestClient

    from backend.app import app

    client = TestClient(app)
    with open(os.path.join(RACINE, "data", "befa_exemple_oivoire.json"), encoding="utf-8") as f:
        dossier = f.read()
    with open(os.path.join(RACINE, "tests", "fixtures", "trends_multiTimeline_huile_de_palme.csv"), "rb") as f:
        csv = f.read()
    r = client.post("/api/befa/generer", data={"dossier": dossier},
                    files=[("fichiers_trends", ("multiTimeline.csv", csv, "text/csv"))])
    assert r.status_code == 200, r.text
    d = r.json()
    assert d["pages"] > 5 and d["sommaire_complet"]
    pdf = client.get(d["pdf"])
    assert pdf.status_code == 200 and pdf.content.startswith(b"%PDF")
    assert 'filename="Dossier_credit_O_IVOIRE.pdf"' in pdf.headers["content-disposition"]
    assert client.get(d["docx"]).content.startswith(b"PK")
    assert client.get("/api/befa/telecharger/inconnu/pdf").status_code == 404

    mauvais = json.loads(dossier)
    mauvais["projet"]["intitule"] = ""
    r = client.post("/api/befa/generer", data={"dossier": json.dumps(mauvais)})
    assert r.status_code == 422 and "intitulé" in r.json()["detail"]
