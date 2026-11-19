"""
Conservation des dossiers sur le serveur (tracabilite, cahier des charges 10.5).

  donnees/dossiers/<id>/dossier.json                   etat complet du dossier (sauvegarde automatique)
  donnees/dossiers/<id>/sources/<horodatage>_<fichier> documents importes (RCCM, statuts, bilans...)
                    .../<horodatage>_<fichier>.extraction.json   resultat de l'extraction
  donnees/dossiers/<id>/documents/<version>/           chaque generation (Business Plan, courriers,
                                                       autres documents) : fichiers + meta.json

Rien n'est jamais ecrase : chaque import et chaque generation cree une
nouvelle entree horodatee, ce qui permet de retrouver les versions
precedentes. Les identifiants sont controles (pas d'acces hors du dossier).
"""
import json
import os
import re
import shutil
import time
import uuid
from datetime import datetime
from typing import Any, Dict, List, Optional

from .referentiels import dossier_donnees

_ID = re.compile(r"^[a-f0-9]{8,32}$")
_VERSION = re.compile(r"^[0-9]{8}-[0-9]{6}-[a-z0-9_]{1,40}$")


class StockageError(ValueError):
    pass


def nouvel_id() -> str:
    return uuid.uuid4().hex[:16]


def _racine(ident: str, creer: bool = False) -> str:
    if not _ID.match(ident or ""):
        raise StockageError("Identifiant de dossier invalide.")
    d = os.path.join(dossier_donnees(), "dossiers", ident)
    if creer:
        os.makedirs(d, exist_ok=True)
    elif not os.path.isdir(d):
        raise StockageError("Dossier introuvable sur le serveur.")
    return d


def _nom_sur(nom: str) -> str:
    base = os.path.basename(nom or "fichier")
    return re.sub(r"[^\w.\-]+", "_", base)[:80] or "fichier"


def _horodatage() -> str:
    return datetime.now().strftime("%Y%m%d-%H%M%S")


def _ecrire_json(chemin: str, data: Any):
    tmp = chemin + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)
    os.replace(tmp, chemin)


def enregistrer_dossier(ident: str, etat: Dict[str, Any]) -> Dict[str, Any]:
    d = _racine(ident, creer=True)
    etat = dict(etat)
    etat["_enregistre_le"] = datetime.now().isoformat(timespec="seconds")
    _ecrire_json(os.path.join(d, "dossier.json"), etat)
    return {"id": ident, "enregistre_le": etat["_enregistre_le"]}


def charger_dossier(ident: str) -> Dict[str, Any]:
    chemin = os.path.join(_racine(ident), "dossier.json")
    if not os.path.exists(chemin):
        raise StockageError("Dossier introuvable sur le serveur.")
    with open(chemin, encoding="utf-8") as f:
        return json.load(f)


def lister_dossiers() -> List[Dict[str, Any]]:
    base = os.path.join(dossier_donnees(), "dossiers")
    if not os.path.isdir(base):
        return []
    res = []
    for ident in os.listdir(base):
        chemin = os.path.join(base, ident, "dossier.json")
        if not (_ID.match(ident) and os.path.exists(chemin)):
            continue
        try:
            with open(chemin, encoding="utf-8") as f:
                etat = json.load(f)
        except (OSError, ValueError):
            continue
        D = etat.get("D", etat)
        e, p = D.get("entreprise") or {}, D.get("projet") or {}
        res.append({"id": ident, "nom": e.get("raison_sociale") or p.get("intitule") or "(sans nom)",
                    "variante": D.get("variante", ""), "projet": p.get("intitule", ""),
                    "enregistre_le": etat.get("_enregistre_le", "")})
    return sorted(res, key=lambda x: x["enregistre_le"], reverse=True)


def supprimer_dossier(ident: str):
    shutil.rmtree(_racine(ident), ignore_errors=True)


def conserver_source(ident: Optional[str], nom: str, contenu: bytes, type_doc: str,
                     extraction: Optional[Dict[str, Any]] = None) -> Optional[str]:
    """Archive un document importe et le resultat de son extraction. Sans identifiant : rien n'est conserve."""
    if not ident:
        return None
    d = os.path.join(_racine(ident, creer=True), "sources")
    os.makedirs(d, exist_ok=True)
    fichier = f"{_horodatage()}_{type_doc}_{_nom_sur(nom)}"
    with open(os.path.join(d, fichier), "wb") as f:
        f.write(contenu)
    if extraction is not None:
        _ecrire_json(os.path.join(d, fichier + ".extraction.json"), extraction)
    return fichier


def lister_sources(ident: str) -> List[Dict[str, Any]]:
    d = os.path.join(_racine(ident), "sources")
    if not os.path.isdir(d):
        return []
    return [{"fichier": f, "taille": os.path.getsize(os.path.join(d, f)),
             "extraction": os.path.exists(os.path.join(d, f + ".extraction.json"))}
            for f in sorted(os.listdir(d), reverse=True) if not f.endswith(".extraction.json")]


def archiver_documents(ident: Optional[str], type_doc: str, fichiers: List[str], meta: Dict[str, Any]) -> Optional[str]:
    """Copie une generation dans une nouvelle version horodatee. Retourne l'identifiant de version."""
    if not ident:
        return None
    type_doc = re.sub(r"[^a-z0-9_]+", "_", type_doc.lower())[:40] or "document"
    version = f"{_horodatage()}-{type_doc}"
    d = os.path.join(_racine(ident, creer=True), "documents", version)
    while os.path.exists(d):  # deux generations dans la meme seconde
        time.sleep(1)
        version = f"{_horodatage()}-{type_doc}"
        d = os.path.join(_racine(ident), "documents", version)
    os.makedirs(d)
    for f in fichiers:
        if f and os.path.exists(f):
            shutil.copy2(f, os.path.join(d, os.path.basename(f)))
    _ecrire_json(os.path.join(d, "meta.json"), {**meta, "type": type_doc, "cree_le": datetime.now().isoformat(timespec="seconds")})
    return version


def lister_versions(ident: str) -> List[Dict[str, Any]]:
    d = os.path.join(_racine(ident), "documents")
    if not os.path.isdir(d):
        return []
    res = []
    for v in sorted(os.listdir(d), reverse=True):
        dv = os.path.join(d, v)
        meta = {}
        if os.path.exists(os.path.join(dv, "meta.json")):
            with open(os.path.join(dv, "meta.json"), encoding="utf-8") as f:
                meta = json.load(f)
        res.append({"version": v, **meta, "fichiers": sorted(x for x in os.listdir(dv) if x != "meta.json")})
    return res


def chemin_document(ident: str, version: str, fichier: str) -> str:
    if not _VERSION.match(version or "") or fichier != os.path.basename(fichier) or fichier == "meta.json":
        raise StockageError("Document introuvable.")
    chemin = os.path.join(_racine(ident), "documents", version, fichier)
    if not os.path.isfile(chemin):
        raise StockageError("Document introuvable.")
    return chemin
