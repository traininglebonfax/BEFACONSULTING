"""
Lot 2 BEFA - import des liasses SYSCOHADA.

Les fichiers tests/fixtures/liasse_*.json sont le resultat REEL de la
lecture par Claude des liasses du dossier "Model Business Plan/" (ETTS 2025,
bilan d'ouverture COREAL 2024). Les tests rejouent la conversion
deterministe et les controles sans aucun appel reseau.
"""
import json
import os

import pytest

from befa import liasse as mod
from befa.liasse import LiasseLue, convertir, selectionner_pages

ICI = os.path.dirname(__file__)
RACINE = os.path.dirname(ICI)
MODELES = os.path.join(RACINE, "Model Business Plan")


def charger(nom: str) -> LiasseLue:
    with open(os.path.join(ICI, "fixtures", nom), encoding="utf-8") as f:
        return LiasseLue.depuis_json(json.load(f))


def test_etts_2025_bilan_et_compte_de_resultat():
    imp = convertir(charger("liasse_etts_2025.json"))
    assert imp.entreprise["sigle"] == "ETTS"
    assert imp.alertes == []  # tous les totaux recalcules concordent avec la liasse

    b = imp.bilan_n
    assert b.exercice == "2025"
    assert b.total_actif == b.total_passif == 203_793_484
    assert b.actif_immobilise_net == 50_000_000
    assert b.creances_clients == 154_937_907
    assert b.capitaux_propres == 146_349_619
    assert b.dettes_fiscales_sociales == 27_443_865
    assert b.tresorerie_passif == 30_000_000
    assert b.tresorerie_nette == -1_144_423 - 30_000_000
    # immobilisations corporelles totalement amorties : brut 12,6 M, net 0
    corporelles = [p for p in b.immobilisations if p.libelle == "Matériel de transport"]
    assert corporelles and corporelles[0].brut == 6_000_000 and corporelles[0].net == 0

    n1 = imp.bilan_n1
    assert n1.exercice == "2024" and not n1.detail_amortissements
    assert n1.total_actif == n1.total_passif == 240_208_623

    assert imp.cr_n.chiffre_affaires == 540_675_290
    assert imp.cr_n.valeur_ajoutee == 72_488_255
    assert imp.cr_n.ebe == 59_623_255
    assert imp.cr_n.resultat_net == 43_955_675
    assert imp.cr_n1.dotations == 12_600_000
    assert imp.cr_n1.resultat_net == 56_123_410


def test_coreal_bilan_ouverture_sans_exercice_precedent():
    imp = convertir(charger("liasse_coreal_ouverture_2024.json"))
    assert imp.bilan_n1 is None and imp.cr_n1 is None
    b = imp.bilan_n
    assert b.total_actif == b.total_passif == 150_365_000
    assert b.capitaux_propres == 7_376_290
    assert imp.cr_n.resultat_net == -2_623_710
    # 143 M d'"Autres dettes" : suggestion de reclassement en compte courant d'associes
    assert len(imp.alertes) == 1 and "compte courant" in imp.alertes[0].lower()


def test_controles_detectent_une_erreur_de_lecture():
    l = charger("liasse_etts_2025.json")
    l.actif["BI"]["net"] += 1_000_000  # l'IA aurait mal lu un chiffre
    alertes = convertir(l).alertes
    assert any("Total actif 2025" in a for a in alertes)
    l = charger("liasse_etts_2025.json")
    l.compte_resultat["RH"]["n"] = -400_000_000
    alertes = convertir(l).alertes
    assert any("valeur ajoutée" in a.lower() for a in alertes)


def test_charges_imprimees_en_positif():
    """Certains logiciels impriment les charges sans signe : la convention
    est retrouvee grace a la valeur ajoutee imprimee (XC)."""
    l = charger("liasse_etts_2025.json")
    for ref, ligne in l.compte_resultat.items():
        if ref in mod.CHARGES:
            for col in ("n", "n1"):
                if ligne.get(col) is not None:
                    ligne[col] = abs(ligne[col])
    imp = convertir(l)
    assert imp.alertes == []
    assert imp.cr_n.valeur_ajoutee == 72_488_255
    assert imp.cr_n.resultat_net == 43_955_675


def test_totaux_utilises_si_sous_lignes_absentes():
    l = LiasseLue.depuis_json({
        "entreprise": {"exercice_clos_le": "31/12/2025"},
        "actif": [
            {"ref": "AI", "libelle": "", "brut": 10_000_000, "amortissements": 4_000_000, "net": 6_000_000, "net_n1": None},
            {"ref": "BK", "libelle": "", "brut": 3_000_000, "amortissements": None, "net": 3_000_000, "net_n1": None},
            {"ref": "BT", "libelle": "", "brut": 1_000_000, "amortissements": None, "net": 1_000_000, "net_n1": None},
            {"ref": "BZ", "libelle": "", "brut": None, "amortissements": None, "net": 10_000_000, "net_n1": None},
        ],
        "passif": [
            {"ref": "CA", "libelle": "", "n": 8_000_000, "n1": None},
            {"ref": "DD", "libelle": "", "n": 2_000_000, "n1": None},
            {"ref": "DZ", "libelle": "", "n": 10_000_000, "n1": None},
        ],
        "compte_resultat": [],
        "remarques": "",
    })
    imp = convertir(l)
    b = imp.bilan_n
    assert b.immobilisations[0].libelle == "Immobilisations corporelles"
    assert b.actif_immobilise_net == 6_000_000
    assert b.dettes_financieres == 2_000_000
    assert b.total_actif == b.total_passif == 10_000_000
    assert imp.alertes == []


def test_lire_liasse_sans_reseau(monkeypatch):
    """La chaine lire_liasse -> convertir avec une reponse IA simulee."""
    with open(os.path.join(ICI, "fixtures", "liasse_etts_2025.json"), encoding="utf-8") as f:
        reponse = json.load(f)
    appels = {}

    def faux_appeler(contenu, **kw):
        appels["schema"] = kw.get("schema")
        appels["blocs"] = [b["type"] for b in contenu]
        return reponse

    monkeypatch.setattr(mod, "appeler", faux_appeler)
    monkeypatch.setattr(mod, "selectionner_pages", lambda pdf: (pdf, [1]))
    l = mod.lire_liasse(b"%PDF-1.4 factice")
    assert appels["blocs"] == ["document", "text"]
    assert appels["schema"]["required"] == ["entreprise", "actif", "passif", "compte_resultat", "remarques"]
    assert convertir(l).bilan_n.total_actif == 203_793_484


def test_lire_liasse_vide_leve_une_erreur(monkeypatch):
    monkeypatch.setattr(mod, "appeler", lambda c, **kw: {"entreprise": {}, "actif": [], "passif": [],
                                                         "compte_resultat": [], "remarques": ""})
    monkeypatch.setattr(mod, "selectionner_pages", lambda pdf: (pdf, [1]))
    with pytest.raises(mod.IaError, match="liasse SYSCOHADA"):
        mod.lire_liasse(b"%PDF")


@pytest.mark.skipif(not os.path.isdir(MODELES), reason="dossier des modeles BEFA absent")
def test_selection_des_pages_sur_les_vraies_liasses():
    for nom in ("Bilan 2025 ETTS  .pdf", "BILAN D'OUVERTURE COREAL FOOD 2024.pdf"):
        with open(os.path.join(MODELES, nom), "rb") as f:
            reduit, pages = selectionner_pages(f.read())
        assert pages == [1, 9, 10, 11]
        assert reduit.startswith(b"%PDF")


# ----------------------------------------------------------------------
# Route HTTP POST /api/befa/liasse
# ----------------------------------------------------------------------
def test_route_import_liasse(monkeypatch):
    from fastapi.testclient import TestClient

    import backend.befa_api as api
    from backend.app import app

    monkeypatch.setattr(api, "lire_liasse", lambda contenu, nom: charger("liasse_etts_2025.json"))
    client = TestClient(app)

    r = client.post("/api/befa/liasse", files={"fichier": ("etts.pdf", b"%PDF-1.4 x", "application/pdf")})
    assert r.status_code == 200, r.text
    d = r.json()
    assert d["controles_ok"] is True
    assert d["bilan_n"]["exercice"] == "2025"
    assert d["apercu"]["colonnes"][:2] == ["2025", "2024"]
    libelles = [l["libelle"] for l in d["apercu"]["structure"]]
    assert libelles == ["Fonds de roulement", "Besoin en fonds de roulement", "Trésorerie nette"]

    r = client.post("/api/befa/liasse", files={"fichier": ("x.txt", b"bonjour", "text/plain")})
    assert r.status_code == 422
