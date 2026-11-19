"""
Referentiels administrables par BEFA Consulting, sans intervention technique :

  - "banques"  : banques et etablissements financiers destinataires des
                 demandes de financement ;
  - "modeles"  : types de documents de la liste « Autres documents a generer ».

Chaque referentiel est un fichier JSON du dossier de donnees (variable
d'environnement BEFA_DONNEES, par defaut `donnees/` a la racine du projet,
non versionne). Au premier usage il est initialise depuis la liste fournie
avec l'application (`data/referentiels/*.json`). L'ecriture est atomique
(fichier temporaire puis renommage) et protegee par un verrou.
"""
import json
import os
import re
import tempfile
import threading
import unicodedata
import uuid
from typing import Any, Dict, List

RACINE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAUTS = os.path.join(RACINE, "data", "referentiels")
NOMS = ("banques", "modeles")
_VERROU = threading.Lock()

CHAMPS = {
    "banques": {"denomination": "", "sigle": "", "type": "Banque", "destinataire": "",
                "titre_destinataire": "Monsieur le Directeur Général", "agence": "", "adresse": "",
                "adresse_agence": "", "boite_postale": "", "ville": "Abidjan", "telephone": "", "email": "",
                "notes": "", "actif": True, "a_verifier": False},
    "modeles": {"titre": "", "categorie": "", "juridique": False, "consigne": "", "calcul": "", "actif": True},
}
OBLIGATOIRES = {"banques": "denomination", "modeles": "titre"}


class ReferentielError(ValueError):
    pass


def dossier_donnees() -> str:
    d = os.environ.get("BEFA_DONNEES") or os.path.join(RACINE, "donnees")
    os.makedirs(d, exist_ok=True)
    return d


def _chemin(nom: str) -> str:
    if nom not in NOMS:
        raise ReferentielError(f"Référentiel inconnu : {nom}.")
    return os.path.join(dossier_donnees(), f"{nom}.json")


def _slug(t: str) -> str:
    t = unicodedata.normalize("NFKD", t).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]+", "_", t).strip("_")[:40] or uuid.uuid4().hex[:8]


def _ecrire(nom: str, elements: List[Dict[str, Any]]):
    chemin = _chemin(nom)
    fd, tmp = tempfile.mkstemp(dir=os.path.dirname(chemin), suffix=".tmp")
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        json.dump(elements, f, ensure_ascii=False, indent=1)
    os.replace(tmp, chemin)


def lister(nom: str, actifs_seulement: bool = False) -> List[Dict[str, Any]]:
    chemin = _chemin(nom)
    with _VERROU:
        if not os.path.exists(chemin):
            with open(os.path.join(DEFAUTS, f"{nom}.json"), encoding="utf-8") as f:
                _ecrire(nom, [_complet(nom, x) for x in json.load(f)])
        with open(chemin, encoding="utf-8") as f:
            elements = json.load(f)
    return [x for x in elements if x.get("actif", True)] if actifs_seulement else elements


def _complet(nom: str, x: Dict[str, Any]) -> Dict[str, Any]:
    res = {**CHAMPS[nom], **{k: v for k, v in x.items() if k in CHAMPS[nom] or k == "id"}}
    if not res.get("id"):
        res["id"] = _slug(res.get(OBLIGATOIRES[nom]) or "")
    return res


def obtenir(nom: str, ident: str) -> Dict[str, Any]:
    for x in lister(nom):
        if x["id"] == ident:
            return x
    raise ReferentielError(f"Élément introuvable : {ident}.")


def enregistrer(nom: str, element: Dict[str, Any]) -> Dict[str, Any]:
    """Ajoute (sans id) ou modifie (avec id existant) un element."""
    if not str(element.get(OBLIGATOIRES[nom]) or "").strip():
        raise ReferentielError(f"Le champ « {OBLIGATOIRES[nom]} » est obligatoire.")
    elements = lister(nom)
    x = _complet(nom, element)
    with _VERROU:
        ids = [e["id"] for e in elements]
        if element.get("id") and element["id"] in ids:
            elements[ids.index(element["id"])] = x
        else:
            base, k = x["id"], 2
            while x["id"] in ids:
                x["id"] = f"{base}_{k}"
                k += 1
            elements.append(x)
        _ecrire(nom, elements)
    return x


def supprimer(nom: str, ident: str):
    elements = lister(nom)
    restants = [x for x in elements if x["id"] != ident]
    if len(restants) == len(elements):
        raise ReferentielError(f"Élément introuvable : {ident}.")
    with _VERROU:
        _ecrire(nom, restants)


def reinitialiser(nom: str):
    with _VERROU:
        if os.path.exists(_chemin(nom)):
            os.remove(_chemin(nom))
    return lister(nom)
