"""
Lot 6 BEFA - services du formulaire web (exemples, apercu financier, page d'accueil).
L'interface elle-meme a ete testee dans Chrome (chargement d'exemple, saisie, apercu
recalcule, relecture de la redaction, generation PDF/Word, erreur de credit API).
"""
from fastapi.testclient import TestClient

from backend.app import app

client = TestClient(app)


def test_pages_servies():
    accueil = client.get("/")
    assert accueil.status_code == 200 and "Business Plan — générateur BEFA" in accueil.text
    assert client.get("/excel.html").status_code == 200  # ancien outil Excel conserve
    logo = client.get("/assets/branding/logo_befa.png")
    assert logo.status_code == 200 and logo.content.startswith(b"\x89PNG")


def test_exemples():
    for nom, variante in (("coreal", "nouveau_projet"), ("oivoire", "entreprise_existante")):
        r = client.get(f"/api/befa/exemples/{nom}")
        assert r.status_code == 200 and r.json()["variante"] == variante
    assert client.get("/api/befa/exemples/inconnu").status_code == 404


def test_apercu_projection():
    dossier = client.get("/api/befa/exemples/coreal").json()
    r = client.post("/api/befa/projection", json={"projection": dossier["projection"], "nom": "COREAL FOOD"})
    d = r.json()
    assert r.status_code == 200 and d["erreurs"] == []
    assert d["synthese"]["cout_investissements"] == 2_700_702_260
    assert d["compte"][2]["libelle"] == "Chiffre d'affaires total"
    assert d["plan"]["colonnes"][0] == "Année 0"
    assert d["commentaires"]["compte_exploitation"][0].startswith("COREAL FOOD escompte")


def test_apercu_projection_saisie_incomplete():
    d = client.post("/api/befa/projection", json={"projection": {"annees": ["A1", "A2"], "produits": []}}).json()
    assert d["erreurs"] == ["Au moins un produit ou service est obligatoire."]
    d = client.post("/api/befa/projection", json={"projection": {
        "annees": ["A1"], "produits": [{"libelle": "X", "montants": [1, 2]}]}}).json()
    assert any("une par année" in e for e in d["erreurs"])
    d = client.post("/api/befa/projection", json={"projection": {"annees": ["A1"], "produits": [{"lib": "X"}]}}).json()
    assert d["erreurs"] and d["erreurs"][0].startswith("Saisie incomplète")
