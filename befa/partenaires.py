"""
Import des principaux clients et fournisseurs depuis un fichier Excel (ou CSV).

Lecture DETERMINISTE (sans IA) : les colonnes sont reconnues par leurs
intitules (synonymes usuels, accents et casse ignores), les parts en
pourcentage sont calculees quand les montants sont disponibles, et les
incoherences sont signalees. Le consultant verifie puis corrige le resultat.

Un classeur peut contenir une feuille "Clients" et une feuille "Fournisseurs" ;
une feuille unique est rattachee au type demande. Un modele vierge est fourni
par `modele_excel()`.
"""
import csv
import io
import re
import unicodedata
from typing import Any, Dict, List, Optional, Tuple

from .extraction import ExtractionError, extension

COLONNES = {
    "clients": {
        "nom": ["nom du client", "client", "nom", "raison sociale", "denomination", "clients"],
        "secteur": ["secteur d'activite", "secteur", "activite", "domaine"],
        "montant": ["chiffre d'affaires realise", "chiffre d'affaires", "ca realise", "ca", "montant",
                    "chiffre d'affaires (fcfa)", "montant facture", "ventes"],
        "part": ["part du chiffre d'affaires", "part dans le chiffre d'affaires", "part du ca", "part", "%",
                 "pourcentage", "part (%)", "poids"],
        "mode_paiement": ["mode de paiement", "mode de reglement", "moyen de paiement", "reglement"],
        "delai": ["delai de paiement", "delai de reglement", "delai"],
        "localisation": ["localisation", "ville", "adresse", "pays"],
    },
    "fournisseurs": {
        "nom": ["nom du fournisseur", "fournisseur", "nom", "raison sociale", "denomination", "fournisseurs"],
        "secteur": ["domaine d'activite", "secteur d'activite", "secteur", "domaine", "activite"],
        "nature": ["nature des biens ou services fournis", "nature des biens", "nature", "biens ou services",
                   "produits fournis", "objet"],
        "montant": ["volume des achats", "montant annuel des transactions", "montant annuel", "achats",
                    "montant", "volume", "montant des achats"],
        "part": ["part dans les achats globaux", "part dans les achats", "part des achats", "part", "%",
                 "pourcentage", "poids"],
        "mode_paiement": ["mode de reglement", "mode de paiement", "reglement"],
        "delai": ["delai de reglement", "delai de paiement", "delai"],
        "localisation": ["localisation", "ville", "adresse", "pays"],
    },
}
OBLIGATOIRE = "nom"
SEUIL_DEPENDANCE = 0.30  # un client > 30 % du CA : dependance commerciale a signaler


def _norm(t: Any) -> str:
    t = unicodedata.normalize("NFKD", str(t or "")).encode("ascii", "ignore").decode().lower()
    # apostrophes et ponctuation ignorees : « Chiffre d'affaires » = « Chiffre d affaires » = « chiffre_d_affaires »
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9%]+", " ", t)).strip()


def _associer(entetes: List[Any], type_: str) -> Dict[str, int]:
    """Colonne -> index. Correspondance exacte d'abord, puis par inclusion du synonyme le plus long."""
    normes = [_norm(e) for e in entetes]
    res: Dict[str, int] = {}
    for passe in ("exacte", "inclusion"):
        for champ, synonymes in COLONNES[type_].items():
            if champ in res:
                continue
            for syn in sorted(synonymes, key=len, reverse=True):
                s = _norm(syn)
                for i, e in enumerate(normes):
                    if i in res.values() or not e:
                        continue
                    if (passe == "exacte" and e == s) or (passe == "inclusion" and len(s) > 2 and s in e):
                        res[champ] = i
                        break
                if champ in res:
                    break
    return res


def _nombre(v: Any) -> Optional[float]:
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(" ", "").replace(" ", "").replace(" ", "")
    pct = s.endswith("%")
    s = s.rstrip("%").replace("FCFA", "").replace("F", "")
    if re.fullmatch(r"-?\d{1,3}(\.\d{3})+(,\d+)?", s):  # 1.250.000,50
        s = s.replace(".", "")
    s = s.replace(",", ".")
    try:
        x = float(s)
    except ValueError:
        return None
    return x / 100 if pct else x


def _feuilles(nom: str, contenu: bytes) -> List[Tuple[str, List[List[Any]]]]:
    ext = extension(nom)
    if ext in ("xlsx", "xlsm"):
        from openpyxl import load_workbook

        try:
            wb = load_workbook(io.BytesIO(contenu), data_only=True, read_only=True)
        except Exception as e:
            raise ExtractionError(f"Classeur Excel illisible ({e}).")
        return [(ws.title, [list(r) for r in ws.iter_rows(values_only=True)]) for ws in wb.worksheets]
    if ext in ("csv", "txt"):
        for enc in ("utf-8-sig", "cp1252"):
            try:
                texte = contenu.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        try:
            dialecte = csv.Sniffer().sniff(texte[:4000], delimiters=";,\t")
        except csv.Error:
            dialecte = csv.excel
        return [("CSV", [r for r in csv.reader(io.StringIO(texte), dialecte)])]
    if ext == "xls":
        raise ExtractionError("Ancien format Excel (.xls) : enregistrez le fichier au format .xlsx.")
    raise ExtractionError("Format non pris en charge : importez un fichier Excel (.xlsx) ou CSV.")


def _type_feuille(titre: str) -> Optional[str]:
    t = _norm(titre)
    if "fournisseur" in t:
        return "fournisseurs"
    if "client" in t:
        return "clients"
    return None


def _lire_feuille(lignes: List[List[Any]], type_: str) -> Tuple[List[Dict[str, Any]], Dict[str, int], List[str]]:
    # ligne d'en-tete : la premiere (parmi les 10 premieres) ou l'on reconnait la colonne "nom"
    for i, l in enumerate(lignes[:10]):
        assoc = _associer(l, type_)
        if OBLIGATOIRE in assoc and len(assoc) >= 2:
            break
    else:
        raise ExtractionError("Colonnes non reconnues : la première ligne doit contenir les intitulés "
                              "(ex. « Nom du client », « Secteur d'activité », « Chiffre d'affaires »...). "
                              "Téléchargez le modèle Excel pour voir le format attendu.")
    non_reconnues = [str(e) for j, e in enumerate(l) if e not in (None, "") and j not in assoc.values()]
    res = []
    for brut in lignes[i + 1:]:
        if not brut or all(v in (None, "") for v in brut):
            continue

        def cel(champ):
            j = assoc.get(champ)
            return brut[j] if j is not None and j < len(brut) else None

        nom = str(cel("nom") or "").strip()
        if not nom or _norm(nom) in ("total", "total general", "sous total"):
            continue
        p = {"nom": nom}
        for champ in ("secteur", "nature", "mode_paiement", "delai", "localisation"):
            if champ in assoc:
                v = cel(champ)
                p[champ] = "" if v is None else str(v).strip()
        p["montant"] = _nombre(cel("montant")) if "montant" in assoc else None
        part = _nombre(cel("part")) if "part" in assoc else None
        if part is not None and part > 1.0001:
            part = part / 100  # saisi en pourcentage entier (25 pour 25 %)
        p["part"] = part
        res.append(p)
    return res, {k: str(l[v]) for k, v in assoc.items()}, non_reconnues


def controler(liste: List[Dict[str, Any]], type_: str, total_reference: Optional[float] = None) -> Tuple[List[str], List[str]]:
    """Calcule les parts manquantes et retourne (alertes, informations)."""
    alertes, infos = [], []
    montants = [p["montant"] for p in liste if p.get("montant") is not None]
    somme = sum(montants)
    base = total_reference or somme
    calculees = 0
    for p in liste:
        if p.get("montant") is not None and p["montant"] < 0:
            alertes.append(f"{p['nom']} : montant négatif.")
        if p.get("part") is None and p.get("montant") and base:
            p["part"] = p["montant"] / base
            p["part_calculee"] = True
            calculees += 1
        elif p.get("part") is not None and p.get("montant") and base:
            ecart = abs(p["part"] - p["montant"] / base)
            if ecart > 0.01:
                alertes.append(f"{p['nom']} : part indiquée {p['part'] * 100:.1f} % alors que le montant représente "
                               f"{p['montant'] / base * 100:.1f} % du total.")
    if calculees:
        ref = "du chiffre d'affaires de l'entreprise" if total_reference and type_ == "clients" else (
            "des achats globaux" if total_reference else "du total importé")
        infos.append(f"{calculees} part(s) calculée(s) automatiquement à partir des montants ({ref}).")
    noms = [_norm(p["nom"]) for p in liste]
    doublons = sorted({p["nom"] for p, n in zip(liste, noms) if noms.count(n) > 1})
    if doublons:
        alertes.append("Doublons : " + ", ".join(doublons) + ".")
    total_parts = sum(p["part"] for p in liste if p.get("part") is not None)
    if total_parts > 1.005:
        alertes.append(f"Les parts totalisent {total_parts * 100:.1f} % : supérieur à 100 %.")
    if total_reference and somme > total_reference * 1.005:
        quoi = "le chiffre d'affaires" if type_ == "clients" else "le total des achats"
        alertes.append(f"La somme des montants ({somme:,.0f} FCFA) dépasse {quoi} de référence "
                       f"({total_reference:,.0f} FCFA).".replace(",", " "))
    if type_ == "clients":
        for p in liste:
            if (p.get("part") or 0) >= SEUIL_DEPENDANCE:
                infos.append(f"Dépendance commerciale : {p['nom']} représente {p['part'] * 100:.0f} % du chiffre "
                             "d'affaires.")
    manque = [p["nom"] for p in liste if p.get("montant") is None and p.get("part") is None]
    if manque and len(manque) < len(liste):
        alertes.append("Montant et part manquants pour : " + ", ".join(manque[:8]) + ("…" if len(manque) > 8 else "") + ".")
    return alertes, infos


def importer(nom: str, contenu: bytes, type_: str, total_reference: Optional[float] = None) -> Dict[str, Any]:
    """type_ : "clients", "fournisseurs" ou "auto" (feuilles nommees Clients / Fournisseurs)."""
    feuilles = _feuilles(nom, contenu)
    resultat: Dict[str, Any] = {"clients": None, "fournisseurs": None, "fichier": nom}
    for titre, lignes in feuilles:
        t = _type_feuille(titre) or (type_ if type_ in COLONNES else None)
        if t is None or resultat[t] is not None or not any(any(v not in (None, "") for v in l) for l in lignes):
            continue
        try:
            liste, assoc, non_reconnues = _lire_feuille(lignes, t)
        except ExtractionError:
            if len(feuilles) > 1 and _type_feuille(titre) is None:
                continue
            raise
        alertes, infos = controler(liste, t, total_reference if t == "clients" else None)
        if "montant" not in assoc and "part" not in assoc:
            infos.append("Aucune colonne de montant ni de part reconnue : les pourcentages ne peuvent pas être calculés "
                         "(intitulés attendus : « Chiffre d'affaires réalisé » / « Montant annuel des achats »).")
        resultat[t] = {"feuille": titre, "lignes": liste, "colonnes": assoc,
                       "colonnes_ignorees": non_reconnues, "alertes": alertes, "infos": infos}
    if resultat["clients"] is None and resultat["fournisseurs"] is None:
        raise ExtractionError("Aucune liste de clients ou de fournisseurs reconnue dans ce fichier.")
    return resultat


def modele_excel() -> bytes:
    """Classeur modele : une feuille Clients, une feuille Fournisseurs."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill

    wb = Workbook()
    entetes = {
        "Clients": ["Nom du client", "Secteur d'activité", "Chiffre d'affaires réalisé (FCFA)",
                    "Part du chiffre d'affaires (%)", "Mode de paiement", "Délai de paiement"],
        "Fournisseurs": ["Nom du fournisseur", "Domaine d'activité", "Nature des biens ou services fournis",
                         "Montant annuel des achats (FCFA)", "Part dans les achats globaux (%)",
                         "Mode de règlement", "Délai de règlement"],
    }
    for k, (titre, cols) in enumerate(entetes.items()):
        ws = wb.active if k == 0 else wb.create_sheet()
        ws.title = titre
        ws.append(cols)
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="F27F24")
        for i, col in enumerate(cols):
            ws.column_dimensions[chr(65 + i)].width = max(16, len(col) + 2)
    sortie = io.BytesIO()
    wb.save(sortie)
    return sortie.getvalue()
