"""
Lot 4 BEFA - redaction.

Fixtures REELLES (appels Claude du 14/11/2026 sur data/befa_exemple_coreal.json) :
  - redaction_coreal_recherche.json : recherche web (48 faits cites, 32 sources) ;
  - redaction_coreal_brut.json      : redaction complete des 10 sections + SWOT.
Les tests rejouent la mise en forme et les controles sans reseau.
"""
import json
import os

import pytest

from befa import commentaires as C
from befa import redaction as R
from befa.dossier import DossierBEFA
from befa.historique import AnalyseHistorique
from befa.ia import IaError
from befa.projection import LigneCharge, calculer_projection
from befa.redaction import Fait, Recherche, Source

ICI = os.path.dirname(__file__)
RACINE = os.path.dirname(ICI)


def dossier() -> DossierBEFA:
    with open(os.path.join(RACINE, "data", "befa_exemple_coreal.json"), encoding="utf-8") as f:
        return DossierBEFA.from_dict(json.load(f))


def fixture(nom):
    with open(os.path.join(ICI, "fixtures", nom), encoding="utf-8") as f:
        return json.load(f)


def recherche_reelle() -> Recherche:
    d = fixture("redaction_coreal_recherche.json")
    return Recherche([Fait(**f) for f in d["faits"]], [Source(**s) for s in d["sources"]],
                     [tuple(x) for x in d["consultees"]])


# ----------------------------------------------------------------------
# Commentaires financiers (programme)
# ----------------------------------------------------------------------
def test_formats_montants():
    assert C.montant(16_084_640_000) == "16,08 milliards FCFA"
    assert C.montant(1_229_208_345) == "1,23 milliard FCFA"
    assert C.montant(182_466_271) == "182 millions FCFA"
    assert C.montant(15_393_200) == "15,4 millions FCFA"
    assert C.montant(-2_623_710) == "−2,6 millions FCFA"
    assert C.montant(250_000) == "250 000 FCFA"
    assert C.en("Année 3") == "en année 3" and C.en("2026") == "en 2026"
    assert C.a_partir("Année 3") == "à partir de l'année 3" and C.fin_de("2029") == "fin 2029"


def test_commentaires_exemple_coreal():
    d = dossier()
    p = calculer_projection(d.projection)
    coms = R.commentaires_financiers(d, p)
    assert coms["compte_exploitation"][0] == (
        "COREAL FOOD escompte un chiffre d'affaires qui s'améliore progressivement, passant de 16,08 milliards "
        "FCFA en année 1 à 24,13 milliards FCFA à partir de l'année 3, niveau maintenu jusqu'à l'année 6.")
    assert coms["financement"][0].startswith("Le coût total des investissements s'élève à 2 700 702 260 "
                                              "FCFA. Il est financé par un apport du promoteur de 829 753 626 "
                                              "FCFA, soit 31 % du coût total")
    assert "remboursements trimestriels" in coms["financement"][1]
    assert coms["couverture_dette"][1].startswith("Sur toute la période de remboursement, ce ratio est supérieur à 1")
    assert "analyse_bilan" not in coms  # nouveau projet : pas d'historique


def test_commentaires_cas_defavorables():
    d = dossier()
    d.projection.charges.append(LigneCharge("Charges exceptionnelles", "services_exterieurs",
                                            montants=[16_000_000_000] + [0] * 5))
    p = calculer_projection(d.projection)
    ce = C.compte_exploitation(p, "X")
    assert any(x.startswith("Le résultat net est déficitaire en année 1") and "dès l'année 2" in x for x in ce)
    pf = C.plan_financement(p, "X")
    assert "négatifs en année 1" in pf[1]


def test_commentaires_historique_oivoire():
    from tests.test_befa_finance import oivoire_analyse

    h = oivoire_analyse()
    b = C.analyse_bilan(h, "O'IVOIRE")
    assert b[0].startswith("Au terme de l'exercice 2025, les emplois stables évoluent de −28 %. Les immobilisations "
                           "brutes étant inchangées")
    assert "passent de 126 millions FCFA en 2024 à 164 millions FCFA en 2025" in b[1]
    assert "Le fonds de roulement (141 millions FCFA) couvre entièrement le besoin en fonds de roulement " \
           "(126 millions FCFA)" in b[2]
    assert b[-1] == "En conclusion, O'IVOIRE présente au 2025 une situation financière saine."
    cr = C.analyse_compte_resultat(h, "O'IVOIRE")
    assert cr[0] == ("O'IVOIRE clôture l'exercice 2025 avec un résultat net de 128 millions FCFA, contre "
                     "102 millions FCFA l'exercice précédent.")
    assert "en hausse de 110 millions FCFA (+26 %)" in cr[1]


# ----------------------------------------------------------------------
# Recherche web : seules les affirmations citees sont retenues
# ----------------------------------------------------------------------
def test_analyse_reponse_recherche():
    cit = lambda url, titre="T", extrait="x": {"type": "web_search_result_location", "url": url, "title": titre,
                                                "cited_text": extrait}
    blocs = [
        {"type": "server_tool_use"},
        {"type": "web_search_tool_result", "content": [{"url": "https://a.ci", "title": "A"}, {"url": "https://b.org"}]},
        {"type": "web_search_tool_result", "content": {"type": "web_search_tool_result_error", "error_code": "max_uses_exceeded"}},
        {"type": "text", "text": "D'après l'USDA, ", "citations": None},
        {"type": "text", "text": "la production ivoirienne atteint 575 000 tonnes en 2024/25", "citations": [cit("https://a.ci", "A", "575,000 MT")]},
        {"type": "text", "text": ".", "citations": None},
        {"type": "text", "text": "Cette phrase n'a aucune source et doit être écartée.", "citations": None},
        {"type": "text", "text": "Les importations progressent de 10 % selon la BAD.", "citations": [cit("https://b.org"), cit("https://a.ci")]},
    ]
    r = R.analyser_reponse_recherche(blocs)
    assert [s.id for s in r.sources] == ["S1", "S2"] and r.sources[0].url == "https://a.ci"
    assert r.faits[0].texte == "D'après l'USDA, la production ivoirienne atteint 575 000 tonnes en 2024/25."
    assert r.faits[0].sources == ["S1"] and r.faits[0].extraits == ["575,000 MT"]
    assert r.faits[1].sources == ["S2", "S1"]
    assert len(r.faits) == 2 and len(r.consultees) == 2


def test_recherche_reelle_enregistree():
    r = recherche_reelle()
    assert len(r.faits) == 48 and len(r.sources) == 32
    assert all(f.sources and all(r.source(s) for s in f.sources) for f in r.faits)
    assert all(s.url.startswith("http") for s in r.sources)


# ----------------------------------------------------------------------
# Controle des chiffres
# ----------------------------------------------------------------------
def test_verifier_chiffres():
    corpus = "CA 16084640000 ; investissement 2700702260 ; 30 tonnes / heure ; baisse de −48 % ; 575 000 tonnes"
    assert R.verifier_chiffres("Le CA atteint 16,08 milliards FCFA [S3] et le coût 2,7 milliards.", corpus) == []
    assert R.verifier_chiffres("Capacité de 30 t/h, recul de 48 % depuis 2021, 575 000 t.", corpus) == []
    assert R.verifier_chiffres("La production mondiale atteint 73 millions de tonnes.", corpus) == ["73 millions"]
    assert R.verifier_chiffres("Trois phases, 4 équipes, 12 mois.", corpus) == ["12"]


# ----------------------------------------------------------------------
# Assemblage de la reponse IA reelle
# ----------------------------------------------------------------------
def test_assemblage_redaction_reelle():
    from befa import trends

    d = dossier()
    with open(os.path.join(ICI, "fixtures", "trends_multiTimeline_huile_de_palme.csv"), "rb") as f:
        t = trends.lire_fichiers([("multiTimeline.csv", f.read())])
    d.trends_faits, d.trends_source = trends.faits(t), trends.source(t)  # comme lors de l'appel reel
    p = calculer_projection(d.projection)
    red = R.assembler(d, fixture("redaction_coreal_brut.json"), recherche_reelle(), R.commentaires_financiers(d, p), p)
    assert [s.titre for s in red.sections][:3] == ["Introduction", "Cadre institutionnel et réglementaire",
                                                   "Justification du projet"]
    assert len(red.sections) == 10 and all(s.paragraphes or s.puces for s in red.sections)
    assert len(red.section("risques").puces) >= 8
    assert all(len(v) >= 4 for v in red.swot.values())
    assert red.alertes == []


def test_assemblage_controles():
    d = dossier()
    rech = Recherche([Fait("La CI produit 575 000 t.", ["S1"])], [Source("S1", "USDA", "https://usda.gov")])
    brut = {
        "sections": [{"cle": "introduction", "paragraphes": [
            "La Côte d'Ivoire produit 575 000 tonnes [S1][S9].", "Le marché pèse 950 milliards FCFA."], "puces": []}],
        "swot": {"forces": ["Site proche des planteurs"], "faiblesses": [], "opportunites": [], "menaces": []},
    }
    red = R.assembler(d, brut, rech, {}, None)
    intro = red.section("introduction")
    assert intro.paragraphes[0] == "La Côte d'Ivoire produit 575 000 tonnes [S1]."
    assert [s.id for s in red.sources] == ["S1"]
    assert any("[S9] inconnu" in a for a in red.alertes)
    assert any("950 milliards" in a for a in red.alertes)
    assert sum("non rédigée" in a for a in red.alertes) == 9  # 9 autres sections manquantes


def test_rediger_sans_reseau(monkeypatch):
    d = dossier()
    p = calculer_projection(d.projection)
    appels = []

    def faux_appeler(contenu, **kw):
        appels.append(kw)
        return fixture("redaction_coreal_brut.json")

    def recherche_ko(_):
        raise IaError("réseau coupé")

    monkeypatch.setattr(R, "appeler", faux_appeler)
    monkeypatch.setattr(R, "rechercher", recherche_ko)
    red = R.rediger(d, p)
    assert red.alertes[0].startswith("Recherche web indisponible (réseau coupé)")
    assert appels[0]["effort"] == "high" and appels[0]["schema"]["required"] == ["sections", "swot"]
    assert len(red.sections) == 10


def test_dossier_validation():
    d = dossier()
    assert d.validate() == []
    d.variante = "entreprise_existante"
    assert any("liasse" in e for e in d.validate())
    d.variante = "autre"
    assert any("Variante inconnue" in e for e in d.validate())


def test_plan_entreprise_existante():
    cles = [c for c, _, _ in R.SECTIONS["entreprise_existante"]]
    assert cles[0] == "presentation_activite" and cles[-1] == "conclusion"
    assert R._schema("entreprise_existante")["properties"]["sections"]["items"]["properties"]["cle"]["enum"] == cles


def test_route_redaction(monkeypatch):
    from fastapi.testclient import TestClient

    from backend.app import app

    monkeypatch.setattr(R, "appeler", lambda contenu, **kw: fixture("redaction_coreal_brut.json"))
    monkeypatch.setattr(R, "rechercher", lambda d: recherche_reelle())
    with open(os.path.join(RACINE, "data", "befa_exemple_coreal.json"), encoding="utf-8") as f:
        corps = json.load(f)
    r = TestClient(app).post("/api/befa/redaction", json=corps)
    assert r.status_code == 200, r.text
    d = r.json()
    assert len(d["sections"]) == 10 and "compte_exploitation" in d["commentaires_financiers"]
    corps["entreprise"]["raison_sociale"] = ""
    r = TestClient(app).post("/api/befa/redaction", json=corps)
    assert r.status_code == 422 and "raison sociale" in r.json()["detail"]
