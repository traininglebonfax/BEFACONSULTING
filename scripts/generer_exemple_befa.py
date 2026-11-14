"""
Genere le Business Plan d'exemple COREAL au format BEFA (Word + PDF), sans
appel IA : redaction et recherche reelles enregistrees dans tests/fixtures,
export Google Trends reel, bilan d'ouverture lu dans la liasse COREAL.

    python -m scripts.generer_exemple_befa [dossier_sortie]
"""
import json
import os
import sys

from befa import trends
from befa.document import generer
from befa.dossier import DossierBEFA
from befa.liasse import LiasseLue, convertir
from befa.projection import calculer_projection
from befa.redaction import Fait, Recherche, Source, assembler, commentaires_financiers

RACINE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FIX = os.path.join(RACINE, "tests", "fixtures")


def charger(nom):
    with open(os.path.join(FIX, nom), encoding="utf-8") as f:
        return json.load(f)


def exemple_coreal():
    with open(os.path.join(RACINE, "data", "befa_exemple_coreal.json"), encoding="utf-8") as f:
        d = DossierBEFA.from_dict(json.load(f))
    imp = convertir(LiasseLue.depuis_json(charger("liasse_coreal_ouverture_2024.json")))
    d.bilan_n, d.cr_n = imp.bilan_n, imp.cr_n
    with open(os.path.join(FIX, "trends_multiTimeline_huile_de_palme.csv"), "rb") as f:
        t = trends.lire_fichiers([("multiTimeline.csv", f.read())])
    d.trends_faits, d.trends_source = trends.faits(t), trends.source(t)
    r = charger("redaction_coreal_recherche.json")
    rech = Recherche([Fait(**x) for x in r["faits"]], [Source(**x) for x in r["sources"]],
                     [tuple(x) for x in r["consultees"]])
    proj = calculer_projection(d.projection)
    red = assembler(d, charger("redaction_coreal_brut.json"), rech, commentaires_financiers(d, proj), proj)
    return d, red, t


if __name__ == "__main__":
    sortie = sys.argv[1] if len(sys.argv) > 1 else os.path.join(RACINE, "sorties")
    d, red, t = exemple_coreal()
    res = generer(d, red, t, sortie, "Business_Plan_COREAL_FOOD_exemple")
    print(f"PDF  : {res.pdf} ({res.pages} pages)\nWord : {res.docx}\nSommaire complet : {res.sommaire_complet}")
    for a in res.alertes + red.alertes:
        print("ALERTE :", a)
