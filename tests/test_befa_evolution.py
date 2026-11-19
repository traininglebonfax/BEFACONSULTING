"""
Evolution « constitution intelligente du dossier de financement » :
profils StartUP / PME-ETI-GE, extraction RCCM + statuts, import Excel des
clients et fournisseurs, trois derniers bilans, bilan d'ouverture,
compte d'exploitation et plan de tresorerie importes, plan O'IVOIRE,
courriers bancaires, autres documents, referentiels, stockage, acces.

Les appels a l'IA sont simules (reponses au format impose) : les tests
verifient tout ce qui est fait par le programme (controles, calculs,
documents). Les essais reels avec l'API sont decrits dans ETAT_PROJET.md.
"""
import base64
import copy
import io
import json
import os
import zipfile

import pytest
from pypdf import PdfReader

from befa import extraction as X
from befa import partenaires as P
from befa.dossier import DossierBEFA, TableauImporte
from befa.projection import calculer_projection
from engine.recalc import RecalcError, find_soffice

RACINE = os.path.dirname(os.path.dirname(__file__))
FIX = os.path.join(RACINE, "tests", "fixtures")
MODELES = os.path.join(RACINE, "Model Business Plan")
RCCM_REEL = os.path.join(MODELES, "RCCM_326402_CI-ABJ-03-2024-B13-07019 (1).pdf")

try:
    find_soffice()
    LIBREOFFICE = True
except RecalcError:
    LIBREOFFICE = False
avec_libreoffice = pytest.mark.skipif(not LIBREOFFICE, reason="LibreOffice absent")


@pytest.fixture(autouse=True)
def donnees_isolees(tmp_path, monkeypatch):
    """Referentiels et dossiers dans un repertoire temporaire (jamais dans donnees/)."""
    monkeypatch.setenv("BEFA_DONNEES", str(tmp_path / "donnees"))
    monkeypatch.delenv("BEFA_MOT_DE_PASSE", raising=False)


def exemple(nom):
    with open(os.path.join(RACINE, "data", f"befa_exemple_{nom}.json"), encoding="utf-8") as f:
        return json.load(f)


def texte_pdf(chemin):
    return "\n".join(p.extract_text() or "" for p in PdfReader(chemin).pages)


def png(largeur=300, hauteur=120, couleur=(242, 127, 36)):
    from PIL import Image

    sortie = io.BytesIO()
    Image.new("RGB", (largeur, hauteur), couleur).save(sortie, "PNG")
    return sortie.getvalue()


def data_url(octets, mime="image/png"):
    return f"data:{mime};base64," + base64.b64encode(octets).decode()


def classeur(feuilles):
    from openpyxl import Workbook

    wb = Workbook()
    for i, (titre, lignes) in enumerate(feuilles.items()):
        ws = wb.active if i == 0 else wb.create_sheet()
        ws.title = titre
        for l in lignes:
            ws.append(l)
    s = io.BytesIO()
    wb.save(s)
    return s.getvalue()


# ----------------------------------------------------------------------
# Profils et regles de gestion
# ----------------------------------------------------------------------
def test_startup_aucun_champ_entreprise_obligatoire():
    d = DossierBEFA.from_dict({"variante": "startup", "projet": {"intitule": "Unité de jus", "secteur": "Agro"}})
    assert d.variante == "nouveau_projet" and d.validate() == []
    assert d.nom == "Unité de jus" and d.raison_sociale == "Unité de jus"


def test_pme_exige_raison_sociale_et_bilan():
    d = DossierBEFA.from_dict({"variante": "pme_eti_ge", "projet": {"intitule": "X", "secteur": "Y"}})
    erreurs = d.validate()
    assert any("raison sociale" in e for e in erreurs) and any("bilan" in e for e in erreurs)


def test_dossier_tolere_les_champs_inconnus_et_lignes_vides():
    d = exemple("oivoire")
    d["champ_futur"] = 1
    d["entreprise"]["cle_inconnue"] = "x"
    d["clients"].append({"nom": "", "mode_paiement": ""})
    dd = DossierBEFA.from_dict(d)
    assert all(c.nom for c in dd.clients)


# ----------------------------------------------------------------------
# RCCM et statuts
# ----------------------------------------------------------------------
def test_reparer_pdf_avec_bom():
    brut = b"\xef\xbb\xbf\xef\xbb\xbf%PDF-1.4 contenu"
    assert X.reparer_pdf(brut).startswith(b"%PDF") and X.est_pdf(brut)


@pytest.mark.skipif(not os.path.exists(RCCM_REEL), reason="RCCM reel absent (document client non versionne)")
def test_rccm_reel_lisible_apres_reparation():
    with open(RCCM_REEL, "rb") as f:
        brut = f.read()
    assert not brut.startswith(b"%PDF")  # le fichier fourni commence par un BOM
    bloc = X.blocs_document("rccm.pdf", brut, "RCCM")[0]
    pdf = base64.b64decode(bloc["source"]["data"])
    assert pdf.startswith(b"%PDF") and "CI-ABJ-03-2024-B13-07019" in PdfReader(io.BytesIO(pdf)).pages[0].extract_text()


def reponse_identite(**surcharges):
    r = {
        "entreprise": {c: "" for c in X.CHAMPS_ENTREPRISE} | {"capital": 1_000_000, "objet_social": ["Conseil"]},
        "dirigeants": [{"nom": "KOUASSI Awa", "fonction": "Gérante", "qualite": "Gérante statutaire", "nationalite": "Ivoirienne",
                        "date_naissance": "", "adresse": "", "document": "Statuts"}],
        "associes": [
            {"nom": "KOUASSI Awa", "nature": "Personne physique", "nationalite": "Ivoirienne", "date_naissance": "",
             "qualite": "", "nombre_parts": 60, "montant": None, "pourcentage": None, "document": "Statuts"},
            {"nom": "DIALLO Ali", "nature": "Personne physique", "nationalite": "Malienne", "date_naissance": "",
             "qualite": "", "nombre_parts": 40, "montant": None, "pourcentage": None, "document": "Statuts"}],
        "valeur_nominale_part": 10_000, "provenance": [{"champ": "raison_sociale", "document": "RCCM"}],
        "informations_absentes": ["Numéro de compte contribuable (NCC)"], "divergences": [], "remarques": "",
    }
    r["entreprise"].update(raison_sociale="LEBONFAX SARL", rccm="CI-ABJ-03-2024-B13-07019", forme_juridique="SARL")
    r.update(surcharges)
    return r


def test_controle_identite_calcule_parts_et_signale_ecarts():
    data, alertes, calculs = X.controler_identite(reponse_identite())
    a, b = data["associes"]
    assert a["montant"] == 600_000 and b["montant"] == 400_000  # parts x valeur nominale
    assert a["pourcentage"] == pytest.approx(0.6) and b["pourcentage"] == pytest.approx(0.4)
    assert alertes == [] and len(calculs) == 4
    data, alertes, _ = X.controler_identite(reponse_identite(valeur_nominale_part=20_000))
    assert any("diffère du capital social" in x for x in alertes)  # 2 000 000 d'apports pour 1 000 000 de capital


def test_extraction_identite_rccm_obligatoire(monkeypatch):
    envoye = {}

    def faux_appeler(contenu, **kw):
        envoye["contenu"], envoye["schema"] = contenu, kw["schema"]
        return reponse_identite()

    monkeypatch.setattr(X, "appeler", faux_appeler)
    with pytest.raises(X.ExtractionError):
        X.extraire_identite([("statuts.pdf", b"%PDF-1.4", "statuts")])
    res = X.extraire_identite([("rccm.pdf", b"\xef\xbb\xbf%PDF-1.4 rccm", "rccm"),
                               ("statuts.docx", _docx("Article 1 - Forme : SARL"), "statuts")])
    assert res["entreprise"]["raison_sociale"] == "LEBONFAX SARL"
    assert envoye["contenu"][0]["type"] == "document"  # PDF repare et envoye tel quel
    assert "Article 1" in envoye["contenu"][1]["text"]  # Word converti en texte
    assert "N'invente" in envoye["contenu"][-1]["text"]
    assert res["informations_absentes"] == ["Numéro de compte contribuable (NCC)"]


def _docx(texte):
    import docx

    d = docx.Document()
    d.add_paragraph(texte)
    s = io.BytesIO()
    d.save(s)
    return s.getvalue()


def test_route_identite(monkeypatch):
    from fastapi.testclient import TestClient

    from backend.app import app

    monkeypatch.setattr(X, "appeler", lambda contenu, **kw: reponse_identite())
    c = TestClient(app)
    assert c.post("/api/befa/identite", files={"statuts": ("s.pdf", b"%PDF-1.4", "application/pdf")}).status_code == 422
    ident = c.post("/api/befa/dossiers").json()["id"]
    r = c.post("/api/befa/identite", files={"rccm": ("rccm.pdf", b"%PDF-1.4 x", "application/pdf")},
               data={"dossier_id": ident})
    assert r.status_code == 200 and r.json()["associes"][0]["pourcentage"] == pytest.approx(0.6)
    sources = c.get(f"/api/befa/dossiers/{ident}/versions").json()["sources"]
    assert len(sources) == 1 and sources[0]["extraction"]  # document source et extraction conserves


# ----------------------------------------------------------------------
# Clients et fournisseurs (Excel)
# ----------------------------------------------------------------------
def test_import_clients_fournisseurs_colonnes_et_parts():
    contenu = classeur({
        "Clients": [["Nom du client", "Secteur d'activité", "CA réalisé (FCFA)", "Observations"],
                    ["Conseil régional", "Administration", 300_000_000, "x"],
                    ["Dorado Ivory", "Mines", "150 000 000", ""],
                    ["Particuliers", "", 50_000_000, ""],
                    ["TOTAL", "", 500_000_000, ""]],
        "Fournisseurs": [["Fournisseur", "Domaine", "Nature des biens ou services fournis", "Montant annuel", "Part (%)"],
                         ["SODIS", "Négoce", "Carburant", 80_000_000, 80],
                         ["CMAO", "BTP", "Ciment", 20_000_000, 20]],
    })
    r = P.importer("liste.xlsx", contenu, "auto", total_reference=537_190_400)
    cl, fo = r["clients"], r["fournisseurs"]
    assert [x["nom"] for x in cl["lignes"]] == ["Conseil régional", "Dorado Ivory", "Particuliers"]  # TOTAL ignore
    assert cl["lignes"][0]["part"] == pytest.approx(300_000_000 / 537_190_400)  # base : CA de l'entreprise
    assert any("Dépendance commerciale" in i for i in cl["infos"]) and cl["colonnes_ignorees"] == ["Observations"]
    assert fo["lignes"][0]["nature"] == "Carburant" and fo["lignes"][0]["part"] == pytest.approx(0.8)
    assert fo["alertes"] == []


def test_import_partenaires_incoherences():
    contenu = classeur({"Feuil1": [["Client", "Montant", "Part du CA"], ["A", 100, 0.9], ["A", 100, 0.5], ["B", -5, None]]})
    r = P.importer("c.xlsx", contenu, "clients")["clients"]
    texte = " ".join(r["alertes"])
    assert "Doublons" in texte and "négatif" in texte and "supérieur à 100 %" in texte


def test_modele_excel_reimportable():
    contenu = P.modele_excel()
    from openpyxl import load_workbook

    wb = load_workbook(io.BytesIO(contenu))
    wb["Clients"].append(["Client X", "Commerce", 1000, None, "Virement", "30 jours"])
    s = io.BytesIO()
    wb.save(s)
    r = P.importer("modele.xlsx", s.getvalue(), "auto")
    assert r["clients"]["lignes"][0]["delai"] == "30 jours" and r["clients"]["lignes"][0]["part"] == 1.0


def test_colonnes_non_reconnues():
    with pytest.raises(X.ExtractionError):
        P.importer("x.xlsx", classeur({"F": [["a", "b"], [1, 2]]}), "clients")


# ----------------------------------------------------------------------
# Bilans : fusion de plusieurs liasses en trois exercices
# ----------------------------------------------------------------------
def _import(annee, avec_n1=True):
    from befa.liasse import LiasseLue, convertir

    with open(os.path.join(FIX, "liasse_etts_2025.json"), encoding="utf-8") as f:
        brut = json.load(f)
    brut["entreprise"]["exercice_clos_le"] = f"31/12/{annee}"
    if not avec_n1:
        for t in ("actif", "passif", "compte_resultat"):
            for l in brut[t]:
                for k in ("n1", "net_n1"):
                    if k in l:
                        l[k] = None
    return convertir(LiasseLue.depuis_json(brut))


def test_fusion_trois_exercices():
    from befa.liasse import fusionner

    f = fusionner([_import(2025), _import(2023)])
    assert f["exercices"] == ["2025", "2024", "2023"]
    assert f["bilans"][2].detail_amortissements  # 2023 vient de la colonne N de sa liasse (prioritaire)
    f = fusionner([_import(2025), _import(2022, avec_n1=False)])
    assert any("non consécutifs" in a for a in f["alertes"])


def test_route_liasses(monkeypatch):
    from fastapi.testclient import TestClient

    import befa.liasse as L
    from backend.app import app

    with open(os.path.join(FIX, "liasse_etts_2025.json"), encoding="utf-8") as f:
        brut = json.load(f)

    def faux(contenu, nom):
        b = copy.deepcopy(brut)
        b["entreprise"]["exercice_clos_le"] = "31/12/2023" if "2023" in nom else "31/12/2025"
        return L.LiasseLue.depuis_json(b)

    monkeypatch.setattr(L, "lire_liasse_fichier", faux)
    r = TestClient(app).post("/api/befa/liasses", files=[("fichiers", ("bilan_2025.pdf", b"%PDF-1.4", "application/pdf")),
                                                          ("fichiers", ("bilan_2023.xlsx", b"PK", "application/octet-stream"))])
    assert r.status_code == 200, r.text
    d = r.json()
    assert d["exercices"] == ["2025", "2024", "2023"] and d["bilan_n2"]["exercice"] == "2023"


# ----------------------------------------------------------------------
# Bilan d'ouverture (startup)
# ----------------------------------------------------------------------
def test_bilan_ouverture_equilibre():
    from befa.ouverture import proposer

    d = DossierBEFA.from_dict(exemple("coreal"))
    r = proposer(d.projection, d.entreprise.capital, stocks_initiaux=5_000_000)
    b = r.bilan
    assert abs(b.total_actif - b.total_passif) < 1
    assert b.capital == d.entreprise.capital and b.compte_courant_associes > 0 and b.dettes_financieres > 0
    assert b.stocks == 5_000_000 and r.hypotheses


def test_bilan_ouverture_besoin_non_couvert():
    from befa.ouverture import proposer
    from befa.projection import ProjectionInput

    p = ProjectionInput.from_dict({"annees": ["Année 1"], "produits": [],
                                   "investissements": [{"libelle": "Machine", "montant": 100, "annee": 0}],
                                   "apports": [{"libelle": "Apport", "montant": 30, "annee": 0}]})
    r = proposer(p)
    assert r.bilan.tresorerie_passif == 70 and abs(r.bilan.total_actif - r.bilan.total_passif) < 1e-6
    assert any("besoin de financement non couvert" in a for a in r.alertes)
    assert proposer(None).alertes  # sans donnees financieres : alerte, pas d'erreur


# ----------------------------------------------------------------------
# Compte d'exploitation et plan de tresorerie importes
# ----------------------------------------------------------------------
def tresorerie(solde_final_m2=21):
    return {"titre": "Plan de trésorerie", "nature": "previsionnel", "periodes": ["Mois 1", "Mois 2"], "remarques": "",
            "lignes": [{"libelle": "Solde initial", "categorie": "solde_initial", "valeurs": [1, 11], "total": None},
                       {"libelle": "Prestations", "categorie": "encaissement", "valeurs": [20, 20], "total": 40},
                       {"libelle": "Charges", "categorie": "decaissement", "valeurs": [10, 10], "total": 25},
                       {"libelle": "Trésorerie nette", "categorie": "solde_final", "valeurs": [11, solde_final_m2], "total": None}]}


def test_controles_plan_tresorerie():
    assert X.controler_tableau(tresorerie(), "plan_tresorerie") == [
        "« Charges » : la somme des périodes (20) diffère du total imprimé (25)."]
    alertes = X.controler_tableau(tresorerie(solde_final_m2=99), "plan_tresorerie")
    assert any("Mois 2 : solde final 99" in a for a in alertes)


def test_extraction_compte_exploitation(monkeypatch):
    ce = {"titre": "Compte d'exploitation 12 mois", "nature": "previsionnel", "periodes": ["M1", "M2"], "remarques": "",
          "lignes": [{"libelle": "Prestations de services", "categorie": "chiffre_affaires", "valeurs": [50, 50], "total": 100},
                     {"libelle": "Carburant", "categorie": "autres_achats", "valeurs": [5, 5], "total": 10},
                     {"libelle": "Location de matériel", "categorie": "services_exterieurs", "valeurs": [12, 12], "total": None},
                     {"libelle": "Amortissements", "categorie": "dotations", "valeurs": [2, 2], "total": 4},
                     {"libelle": "Résultat", "categorie": "sous_total", "valeurs": [31, 31], "total": 62}]}
    monkeypatch.setattr(X, "appeler", lambda contenu, **kw: copy.deepcopy(ce))
    xlsx = classeur({"CE": [["Rubrique", "M1", "M2"], ["Prestations", 50, 50]]})
    r = X.extraire_tableau("ce.xlsx", xlsx, "compte_exploitation")
    assert r["alertes"] == [] and r["lignes"][-1]["style"] == "sous_total"
    p = r["proposition_projection"]
    assert p["produits"] == [{"libelle": "Prestations de services", "montant_annuel": 100}]
    assert {"libelle": "Location de matériel", "rubrique": "services_exterieurs", "montant_annuel": 24} in p["charges"]
    assert p["autres"] == {"dotations": 4}


def test_formats_refuses():
    with pytest.raises(X.ExtractionError):
        X.blocs_document("vieux.xls", b"\xd0\xcf\x11\xe0", "x")
    with pytest.raises(X.ExtractionError):
        X.blocs_document("a.exe", b"MZ", "x")


# ----------------------------------------------------------------------
# Business Plan au plan O'IVOIRE
# ----------------------------------------------------------------------
def dossier_pme_complet():
    d = DossierBEFA.from_dict(exemple("oivoire"))
    imp25, imp23 = _import(2025), _import(2023)
    d.bilan_n, d.bilan_n1, d.cr_n, d.cr_n1 = imp25.bilan_n, imp25.bilan_n1, imp25.cr_n, imp25.cr_n1
    d.bilan_n2, d.cr_n2 = imp23.bilan_n, imp23.cr_n
    d.clients[0].montant, d.clients[0].part = 300_000_000, 0.56
    d.identite_visuelle.logo = data_url(png())
    d.compte_exploitation_importe = TableauImporte.from_dict(
        {"titre": "Compte d'exploitation prévisionnel sur les 12 prochains mois", "fichier": "ce.xlsx",
         "periodes": [f"Mois {i}" for i in range(1, 13)],
         "lignes": [{"libelle": "Prestations de services", "categorie": "chiffre_affaires", "valeurs": [50_000_000] * 12,
                     "total": 600_000_000}]})
    return d


@avec_libreoffice
def test_business_plan_pme_plan_oivoire(tmp_path):
    from befa import document as Doc
    from befa.redaction import Redaction, SectionRedigee

    d = dossier_pme_complet()
    red = Redaction("entreprise_existante", [SectionRedigee("impact_besoin", "Impact", ["Texte."],
                                                            ["Réduction des coûts : fin de la location"])],
                    {"forces": ["Demande structurelle élevée : volume d'activité soutenu"], "faiblesses": [],
                     "opportunites": [], "menaces": ["Concurrence informelle : pression sur les prix"]}, {}, [], [], [], [])
    res = Doc.generer(d, red, None, str(tmp_path), "pme")
    assert res.sommaire_complet, res.alertes
    tout = texte_pdf(res.pdf)
    for attendu in ("PREMIÈRE PARTIE : PRÉSENTATION", "a) Historique de la création", "b) Géographie de l'actionnariat",
                    "III. POOL BANCAIRE", "DEUXIÈME PARTIE : ÉTUDE FINANCIÈRE", "I. ANALYSE DU BILAN 2025",
                    "III. ÉVOLUTION SUR LES TROIS DERNIERS EXERCICES", "TCAM", "f) Caractéristiques du besoin",
                    "d) Mode opératoire de décaissement", "b) Source secondaire", "D. ANALYSE DE LA CONCURRENCE",
                    "Axes stratégiques", "Impacts sur le projet", "volume d'activité soutenu", "Dépendance",
                    "ANNEXES", "Mois 12", "Échéancier annuel des emprunts"):
        assert attendu.lower() in tout.lower(), attendu
    assert "DOSSIER DE CRÉDIT" in PdfReader(res.pdf).pages[0].extract_text()
    import docx

    couverture = docx.Document(res.docx)
    assert len(couverture.inline_shapes) >= 2  # logo du client + logo BEFA sur la page de garde
    paysage = [s for s in couverture.sections if s.page_width > s.page_height]
    assert paysage  # annexes a 12 colonnes en paysage


@avec_libreoffice
def test_business_plan_startup_sans_entreprise(tmp_path):
    from befa import document as Doc
    from befa.ouverture import proposer

    d = DossierBEFA.from_dict(exemple("coreal"))
    d.entreprise = DossierBEFA.from_dict({"variante": "startup", "projet": {"intitule": "x", "secteur": "y"}}).entreprise
    d.actionnaires, d.dirigeants, d.clients, d.fournisseurs = [], [], [], []
    d.bilan_ouverture = proposer(d.projection, 10_000_000).bilan
    d.bilan_ouverture_valide = True
    assert d.validate() == []
    res = Doc.generer(d, None, None, str(tmp_path), "startup")
    tout = texte_pdf(res.pdf)
    assert "BUSINESS PLAN" in PdfReader(res.pdf).pages[0].extract_text()
    for attendu in ("a) Genèse du projet et présentation du promoteur", "I. BILAN D'OUVERTURE PRÉVISIONNEL",
                    "estimation et non d'un bilan comptable certifié", "PRÉVISIONS", "Compte d'exploitation prévisionnel"):
        assert attendu in tout, attendu
    assert "Bilan d'ouverture au" not in tout  # aucune donnee presentee comme historique


# ----------------------------------------------------------------------
# Courriers bancaires
# ----------------------------------------------------------------------
def test_courriers_un_par_banque_personnalise():
    from befa import courriers, referentiels

    d = DossierBEFA.from_dict(exemple("oivoire"))
    p = calculer_projection(d.projection)
    b1, b2 = referentiels.lister("banques")[:2]
    c1, c2 = courriers.preparer(d, p, b1), courriers.preparer(d, p, b2)
    assert c1["banque"] != c2["banque"] and b1["sigle"] in " ".join(c1["corps"]) and b1["sigle"] not in " ".join(c2["corps"])
    texte = " ".join(c1["corps"])
    assert "122 475 094 FCFA" in c1["objet"].replace(" ", " ") and "O'IVOIRE SARL" in texte and "Business Plan" in texte
    assert "537 millions FCFA" in texte  # dernier exercice (historique PME)
    s = courriers.preparer(DossierBEFA.from_dict(exemple("coreal")), None, b1)
    assert "accompagnement financier" in s["objet"]  # sans montant : demande generique


@avec_libreoffice
@pytest.mark.parametrize("papier", ["aucun", "image_pleine_page", "bandeau", "docx"])
def test_generation_courriers_et_zip(tmp_path, papier):
    from befa import courriers, referentiels

    d = DossierBEFA.from_dict(exemple("oivoire"))
    if papier == "image_pleine_page":
        d.identite_visuelle.entete = data_url(png(210, 297, (255, 255, 255)))
    elif papier == "bandeau":
        d.identite_visuelle.entete = data_url(png(800, 120))
    elif papier == "docx":
        import docx

        g = docx.Document()
        g.sections[0].header.paragraphs[0].text = "EN-TETE OFFICIEL O'IVOIRE"
        g.add_paragraph("contenu a effacer")
        s = io.BytesIO()
        g.save(s)
        d.identite_visuelle.entete = data_url(s.getvalue(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    banques = referentiels.lister("banques")[:3]
    lettres = [courriers.preparer(d, calculer_projection(d.projection), b) for b in banques]
    lettres[0]["corps"][1] = "Paragraphe modifié par le consultant."
    res = courriers.generer(d, lettres, str(tmp_path))
    assert len(res["courriers"]) == 3 and res["alertes"] == []
    t0 = texte_pdf(res["courriers"][0]["pdf"])
    assert "Paragraphe modifié par le consultant." in t0 and banques[0]["denomination"] in t0
    assert banques[1]["denomination"] not in t0  # jamais un courrier adresse a plusieurs banques
    if papier == "docx":
        assert "EN-TETE OFFICIEL" in t0 and "contenu a effacer" not in t0
    if papier == "aucun":
        assert "RCCM : CI-DAL-01-2025-M-00962" in t0  # mentions legales en pied de page
    with zipfile.ZipFile(res["zip"]) as z:
        assert len(z.namelist()) == 6


def test_routes_courriers(monkeypatch):
    from fastapi.testclient import TestClient

    import befa.courriers as Co
    from backend.app import app

    c = TestClient(app)
    ids = [b["id"] for b in c.get("/api/befa/referentiels").json()["banques"][:2]]
    assert c.post("/api/befa/courriers/preparer", json={"dossier": exemple("oivoire"), "banques": []}).status_code == 422
    r = c.post("/api/befa/courriers/preparer", json={"dossier": exemple("oivoire"), "banques": ids})
    assert r.status_code == 200 and len(r.json()["courriers"]) == 2
    monkeypatch.setattr(Co, "docx_vers_pdf", lambda docx, sortie: docx.replace(".docx", ".pdf"))
    monkeypatch.setattr(Co.zipfile, "ZipFile", _FauxZip)
    ident = c.post("/api/befa/dossiers").json()["id"]
    r = c.post("/api/befa/courriers/generer", json={"dossier": exemple("oivoire"), "courriers": r.json()["courriers"],
                                                     "dossier_id": ident})
    assert r.status_code == 200, r.text
    assert r.json()["version"] and c.get(r.json()["courriers"][0]["docx"]).content.startswith(b"PK")


class _FauxZip:
    def __init__(self, chemin, *a, **k):
        open(chemin, "wb").close()

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    def write(self, *a):
        pass


# ----------------------------------------------------------------------
# Autres documents
# ----------------------------------------------------------------------
def test_autres_documents_juridique_et_calcul(monkeypatch):
    from befa import autres_documents as A
    from befa import referentiels

    monkeypatch.setattr(A, "appeler", lambda contenu, **kw: {
        "titre": "Contrat de maintenance", "informations_manquantes": ["Prix mensuel"],
        "sections": [{"titre": "Article 1 — Objet", "paragraphes": ["Le prestataire [À compléter : nom]..."], "puces": []}]})
    d = DossierBEFA.from_dict(exemple("oivoire"))
    p = calculer_projection(d.projection)
    m = next(x for x in referentiels.lister("modeles") if x["titre"] == "Contrat de maintenance")
    c = A.preparer(d, p, m)
    assert c["juridique"] and "professionnel compétent" in c["avertissement"] and c["tableaux"] == []
    e = next(x for x in referentiels.lister("modeles") if x["calcul"] == "echeancier")
    c = A.preparer(d, p, e)
    tab = c["tableaux"][0]
    assert not c["juridique"] and tab["entetes"][0] == "N°" and len(tab["lignes"]) == 36  # 36 echeances mensuelles
    assert abs(sum(l[4] for l in tab["lignes"]) - d.projection.prets[0].montant) < 2  # capital rembourse = montant


@avec_libreoffice
def test_generation_autre_document(tmp_path):
    from befa import autres_documents as A

    d = DossierBEFA.from_dict(exemple("oivoire"))
    contenu = {"titre": "Contrat de prestation de services", "juridique": True, "avertissement": A.AVERTISSEMENT_JURIDIQUE,
               "sections": [{"titre": "Article 1", "paragraphes": ["Texte modifié."], "puces": ["a"]}],
               "tableaux": A.tableaux_calcules(d, calculer_projection(d.projection), "budget")}
    res = A.generer(d, contenu, str(tmp_path))
    t = texte_pdf(res["pdf"])
    assert "CONTRAT DE PRESTATION DE SERVICES" in t and "PROJET" in t and "Texte modifié." in t
    assert "Compte d'exploitation prévisionnel" in t


# ----------------------------------------------------------------------
# Referentiels, stockage, securite
# ----------------------------------------------------------------------
def test_referentiels_administrables():
    from befa import referentiels as R

    assert len(R.lister("banques")) >= 20 and all(b["a_verifier"] for b in R.lister("banques"))
    x = R.enregistrer("banques", {"denomination": "Banque Test CI", "adresse": "Plateau", "destinataire": "M. X"})
    assert R.obtenir("banques", x["id"])["adresse"] == "Plateau"
    x["adresse"] = "Cocody"
    R.enregistrer("banques", x)
    assert R.obtenir("banques", x["id"])["adresse"] == "Cocody"
    R.supprimer("banques", x["id"])
    with pytest.raises(R.ReferentielError):
        R.obtenir("banques", x["id"])
    with pytest.raises(R.ReferentielError):
        R.enregistrer("modeles", {"titre": ""})
    m = R.enregistrer("modeles", {"titre": "Bail commercial", "categorie": "B. Contrats commerciaux", "juridique": True})
    assert any(e["id"] == m["id"] for e in R.lister("modeles", actifs_seulement=True))
    assert len(R.reinitialiser("modeles")) == 30


def test_stockage_versions_et_securite_chemins(tmp_path):
    from befa import stockage as S

    ident = S.nouvel_id()
    S.enregistrer_dossier(ident, {"D": {"entreprise": {"raison_sociale": "ABC"}, "variante": "nouveau_projet", "projet": {}}})
    assert S.lister_dossiers()[0]["nom"] == "ABC"
    f = tmp_path / "bp.pdf"
    f.write_bytes(b"%PDF")
    v1 = S.archiver_documents(ident, "business_plan", [str(f)], {"pages": 1})
    v2 = S.archiver_documents(ident, "business_plan", [str(f)], {"pages": 2})
    assert v1 != v2 and len(S.lister_versions(ident)) == 2  # rien n'est ecrase
    assert os.path.exists(S.chemin_document(ident, v1, "bp.pdf"))
    for mauvais in ("../../etc", "x/../y"):
        with pytest.raises(S.StockageError):
            S.charger_dossier(mauvais)
    with pytest.raises(S.StockageError):
        S.chemin_document(ident, v1, "../dossier.json")
    assert S.conserver_source(None, "a.pdf", b"x", "rccm") is None  # sans dossier : rien n'est conserve


def test_acces_protege(monkeypatch):
    from fastapi.testclient import TestClient

    from backend.app import app

    c = TestClient(app)
    assert c.get("/api/befa/referentiels").status_code == 200  # sans mot de passe configure : acces libre
    monkeypatch.setenv("BEFA_MOT_DE_PASSE", "s3cret")
    assert c.get("/api/befa/referentiels").status_code == 401
    assert c.get("/api/befa/referentiels", auth=("befa", "faux")).status_code == 401
    assert c.get("/api/befa/referentiels", auth=("befa", "s3cret")).status_code == 200
    assert c.get("/api/health").status_code == 200


def test_secteurs():
    from befa import secteurs

    assert len(secteurs.SECTEURS) == 24 and secteurs.SECTEURS[-1]["libelle"] == "Autres secteurs"
    assert "Indicateurs" in secteurs.reperes("btp")
    assert secteurs.trouver(libelle="Projet : BTP, construction et travaux publics")["code"] == "btp"
    assert secteurs.reperes("inconnu") == ""


def test_entetes_sans_apostrophe_ni_accent():
    contenu = classeur({"Clients": [["NOM DU CLIENT", "Secteur d activite", "Chiffre d affaires realise (FCFA)"],
                                    ["A", "x", 300], ["B", "y", 100]]})
    r = P.importer("c.xlsx", contenu, "auto")["clients"]
    assert r["colonnes"]["montant"].startswith("Chiffre") and r["lignes"][0]["part"] == pytest.approx(0.75)


def test_pages_html_non_mises_en_cache():
    from fastapi.testclient import TestClient

    from backend.app import app

    r = TestClient(app).get("/")
    assert r.status_code == 200 and "no-cache" in r.headers.get("cache-control", "")


def test_fusion_liasses_entreprises_differentes():
    from befa.liasse import fusionner

    a, b = _import(2025), _import(2023)
    b.entreprise = dict(b.entreprise, raison_sociale="AUTRE SOCIETE", ncc="9999999X")
    assert any("entreprises différentes" in x for x in fusionner([a, b])["alertes"])
    assert not any("entreprises différentes" in x for x in fusionner([a, _import(2023)])["alertes"])
