"""
Lot 3 BEFA - exports Google Trends.

tests/fixtures/trends_multiTimeline_huile_de_palme.csv est un export REEL
(interface en francais, "Huile de palme", Cote d'Ivoire, 5 ans, par semaine).
Les autres formats (anglais, regions, requetes associees) reprennent la
structure des exports Google Trends.
"""
import os
from datetime import date, timedelta

import pytest

from befa import graphiques, trends
from befa.trends import Associes, Evolution, InteretRegional, TrendsError, analyser_serie, lire_fichier

ICI = os.path.dirname(__file__)
REEL = os.path.join(ICI, "fixtures", "trends_multiTimeline_huile_de_palme.csv")


def reel() -> bytes:
    with open(REEL, "rb") as f:
        return f.read()


def test_export_reel_francais():
    ev = lire_fichier(reel(), "multiTimeline.csv")
    assert isinstance(ev, Evolution)
    assert ev.categorie == "Toutes catégories" and ev.granularite == "semaine"
    assert ev.zone == "Côte d’Ivoire"
    assert [s.terme for s in ev.series] == ["Huile de palme"]
    assert len(ev.series[0].points) == 261
    assert (ev.debut, ev.fin) == (date(2021, 11, 14), date(2026, 11, 8))

    a = analyser_serie(ev.series[0])
    assert a.tendance == "en baisse"
    assert a.evolution == pytest.approx(-0.479, abs=0.002)
    assert a.pic_date == date(2023, 10, 22) and a.pic_valeur == 100
    assert not a.fiable and a.part_zeros == pytest.approx(36 / 261)
    assert a.mois_forts[0] == 3 and a.mois_faibles[0] == 12  # mars le plus fort, decembre le plus faible
    assert sum(a.saisonnalite) == pytest.approx(12)


def test_faits_et_source_sur_export_reel():
    d = trends.lire_fichiers([("multiTimeline.csv", reel())])
    f = trends.faits(d)
    assert f[0].startswith("Sur la période novembre 2021 – novembre 2026, l'intérêt de recherche pour « Huile de palme » en Côte d’Ivoire est en baisse")
    assert "−48 %" in f[0]
    assert any(x.startswith("Prudence : 14 % des relevés valent 0") for x in f)
    assert any("mars, mai et février" in x for x in f)
    assert trends.source(d) == ("Source : Google Trends, Côte d’Ivoire, novembre 2021 – novembre 2026, "
                                "données exportées par BEFA Consulting.")


def _csv_hebdo(entetes, lignes, cat="Category: All categories", premiere="Week") -> bytes:
    out = [cat, "", ",".join([premiere] + entetes)]
    out += [",".join(map(str, l)) for l in lignes]
    return "\n".join(out).encode("utf-8")


def test_export_anglais_plusieurs_termes_et_valeurs_inferieures_a_1():
    debut = date(2024, 1, 7)
    lignes = [((debut + timedelta(weeks=i)).isoformat(), 20 + i % 10, "<1" if i % 2 else 3) for i in range(104)]
    ev = lire_fichier(_csv_hebdo(["palm oil: (Côte d’Ivoire)", "shea butter: (Côte d’Ivoire)"], lignes))
    assert [s.terme for s in ev.series] == ["palm oil", "shea butter"]
    assert ev.series[1].points[1][1] == 0.5  # "<1"
    f = trends.faits(trends.DonneesTrends(evolution=ev))
    assert any(x.startswith("Popularité relative sur la période (indice moyen) : « palm oil »") for x in f)


def test_export_mensuel_sans_saisonnalite_si_moins_de_deux_ans():
    lignes = [(f"2025-{m:02d}", 50 + m) for m in range(1, 13)]
    ev = lire_fichier(_csv_hebdo(["huile: (Côte d’Ivoire)"], lignes, cat="Catégorie : Toutes catégories", premiere="Mois"))
    assert ev.granularite == "mois"
    a = analyser_serie(ev.series[0])
    assert a.saisonnalite is None and a.mois_forts == []


def test_tendance_en_hausse():
    debut = date(2022, 1, 2)
    lignes = [((debut + timedelta(weeks=i)).isoformat(), 20 + i // 4) for i in range(156)]
    a = analyser_serie(lire_fichier(_csv_hebdo(["x: (CI)"], lignes)).series[0])
    assert a.tendance == "en hausse" and a.pente_annuelle > 0 and a.fiable


def test_interet_par_region():
    texte = ("Catégorie : Toutes catégories\n\nRégion,Huile de palme: (14/11/2021 - 14/11/2026)\n"
             "Lagunes,74\nAbidjan,100\nBas-Sassandra,\nGôh-Djiboua,52\n")
    r = lire_fichier(texte.encode("utf-8"), "geoMap.csv")
    assert isinstance(r, InteretRegional)
    assert r.regions == [("Abidjan", 100), ("Lagunes", 74), ("Gôh-Djiboua", 52)]  # region sans valeur ignoree
    f = trends.faits(trends.DonneesTrends(regions=[r]))
    assert f == ["Régions où « Huile de palme » suscite le plus d'intérêt : Abidjan (100), Lagunes (74), Gôh-Djiboua (52)."]


def test_requetes_associees_francais_et_anglais():
    fr = ("Catégorie : Toutes catégories\n\nPRINCIPALES\nhuile de palme prix,100\nhuile rouge,45\n\n"
          "HAUSSE\nprix huile de palme 2026,Record\nhuile de palme bio,+250 %\n").encode("utf-8")
    a = lire_fichier(fr, "relatedQueries.csv")
    assert isinstance(a, Associes) and a.nature == "requêtes"
    assert a.principales == [("huile de palme prix", "100"), ("huile rouge", "45")]
    assert a.en_hausse[1] == ("huile de palme bio", "+250 %")
    en = "Category: All categories\n\nTOP\npalm oil,100\n\nRISING\npalm oil price,Breakout\n".encode("utf-8")
    b = lire_fichier(en, "relatedEntities.csv")
    assert b.nature == "sujets" and b.en_hausse == [("palm oil price", "Breakout")]


def test_encodage_windows_et_point_virgule():
    texte = "Catégorie : Toutes catégories\n\nSemaine;Attiéké: (Côte d’Ivoire)\n2025-01-05;40\n2025-01-12;42\n"
    ev = lire_fichier(texte.encode("cp1252", errors="replace"), "x.csv")
    assert ev.series[0].terme.startswith("Atti") and len(ev.series[0].points) == 2


def test_fichier_non_reconnu():
    with pytest.raises(TrendsError, match="format non reconnu"):
        lire_fichier(b"nom,prenom\nKoffi,Ama\n", "clients.csv")
    with pytest.raises(TrendsError, match="vide"):
        lire_fichier(b"", "vide.csv")


def test_graphiques_png():
    ev = lire_fichier(reel())
    a = analyser_serie(ev.series[0])
    for png in (graphiques.evolution_interet([(s.terme, s.points) for s in ev.series], ev.granularite),
                graphiques.saisonnalite(a.saisonnalite),
                graphiques.barres_horizontales(["Abidjan", "Lagunes"], [100, 74]),
                graphiques.courbes_annuelles(["A1", "A2"], [("CA", [1e9, 2e9]), ("RN", [1e8, 3e8])])):
        assert png.startswith(b"\x89PNG") and len(png) > 5_000


def test_route_trends():
    from fastapi.testclient import TestClient

    from backend.app import app

    client = TestClient(app)
    geo = "Région,Huile de palme: (CI)\nAbidjan,100\nLagunes,74\n".encode("utf-8")
    r = client.post("/api/befa/trends", files=[("fichiers", ("multiTimeline.csv", reel(), "text/csv")),
                                               ("fichiers", ("geoMap.csv", geo, "text/csv"))])
    assert r.status_code == 200, r.text
    d = r.json()
    assert d["series"][0]["tendance"] == "en baisse"
    assert d["regions"][0]["regions"][0] == ["Abidjan", 100]
    assert set(d["graphiques"]) == {"evolution", "saisonnalite", "regions"}
    assert d["graphiques"]["evolution"].startswith("data:image/png;base64,")
    assert d["avertissement"].startswith("Google Trends mesure un intérêt de recherche relatif")

    r = client.post("/api/befa/trends", files=[("fichiers", ("x.csv", b"nom,prenom\nA,B\n", "text/csv"))])
    assert r.status_code == 422 and "format non reconnu" in r.json()["detail"]
