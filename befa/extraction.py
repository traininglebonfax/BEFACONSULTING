"""
Extraction automatique d'informations depuis les documents importes.

  - `extraire_identite` : RCCM (obligatoire pour une PME / ETI / GE) et statuts
    -> identification, dirigeants, associes et repartition du capital ;
  - `extraire_tableau`  : compte d'exploitation ou plan de tresorerie (PDF,
    Excel, image) -> tableau par periode, rubriques classees ;
  - `blocs_document`    : conversion d'un fichier importe en blocs lisibles
    par Claude (PDF, image, Word, Excel, CSV, texte).

Principes (cahier des charges, section 10.2) :
  - l'IA RECOPIE ce qui est ecrit ; une information absente reste vide et est
    listee dans "informations_absentes" ; jamais de valeur supposee ;
  - les controles (sommes, pourcentages, totaux) sont faits par le programme ;
  - le resultat est un BROUILLON que le consultant verifie et corrige.
"""
import base64
import csv
import io
import re
from typing import Any, Dict, List, Optional, Tuple

from .ia import IaError, appeler, bloc_pdf

TAILLE_MAX = 30 * 1024 * 1024

TYPES_IMAGES = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg", "webp": "image/webp",
                "gif": "image/gif"}


class ExtractionError(ValueError):
    """Fichier illisible ou format non pris en charge (message destine au consultant)."""


# ----------------------------------------------------------------------
# Lecture des fichiers
# ----------------------------------------------------------------------
def reparer_pdf(contenu: bytes) -> bytes:
    """Certains PDF (ex. RCCM telecharges du guichet unique) commencent par un BOM ou des octets parasites
    avant l'en-tete %PDF : on les retire (sinon le fichier est refuse)."""
    i = contenu.find(b"%PDF", 0, 2048)
    return contenu[i:] if i > 0 else contenu


def extension(nom: str) -> str:
    return (nom or "").rsplit(".", 1)[-1].lower() if "." in (nom or "") else ""


def est_pdf(contenu: bytes) -> bool:
    return b"%PDF" in contenu[:2048]


def _texte_excel(contenu: bytes, max_lignes: int = 400) -> str:
    from openpyxl import load_workbook

    try:
        wb = load_workbook(io.BytesIO(contenu), data_only=True, read_only=True)
    except Exception as e:  # fichier corrompu ou .xls ancien format
        raise ExtractionError(f"Classeur Excel illisible ({e}). Enregistrez-le au format .xlsx puis réessayez.")
    morceaux = []
    for ws in wb.worksheets:
        lignes = []
        for row in ws.iter_rows(values_only=True):
            if row is None or all(v in (None, "") for v in row):
                continue
            lignes.append(" | ".join("" if v is None else str(v).strip() for v in row).rstrip(" |"))
            if len(lignes) >= max_lignes:
                break
        if lignes:
            morceaux.append(f"### Feuille « {ws.title} »\n" + "\n".join(lignes))
    if not morceaux:
        raise ExtractionError("Le classeur Excel est vide.")
    return "\n\n".join(morceaux)


def _texte_docx(contenu: bytes) -> str:
    import docx

    try:
        doc = docx.Document(io.BytesIO(contenu))
    except Exception as e:
        raise ExtractionError(f"Document Word illisible ({e}).")
    parties = [p.text for p in doc.paragraphs if p.text.strip()]
    for t in doc.tables:
        for r in t.rows:
            parties.append(" | ".join(c.text.strip() for c in r.cells))
    return "\n".join(parties)


def _texte_csv(contenu: bytes) -> str:
    for enc in ("utf-8-sig", "utf-16", "cp1252"):
        try:
            texte = contenu.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ExtractionError("Fichier texte illisible (encodage inconnu).")
    try:
        dialecte = csv.Sniffer().sniff(texte[:4000], delimiters=";,\t")
        return "\n".join(" | ".join(r) for r in csv.reader(io.StringIO(texte), dialecte))
    except csv.Error:
        return texte


def blocs_document(nom: str, contenu: bytes, titre: str = "") -> List[Dict[str, Any]]:
    """Blocs de contenu Claude pour un fichier importe."""
    if len(contenu) > TAILLE_MAX:
        raise ExtractionError(f"{nom} : fichier trop volumineux (30 Mo maximum).")
    if not contenu:
        raise ExtractionError(f"{nom} : fichier vide.")
    ext = extension(nom)
    titre = titre or nom
    if est_pdf(contenu):
        return [bloc_pdf(reparer_pdf(contenu), titre)]
    if ext in TYPES_IMAGES:
        return [{"type": "text", "text": f"Document « {titre} » (image) :"},
                {"type": "image", "source": {"type": "base64", "media_type": TYPES_IMAGES[ext],
                                             "data": base64.standard_b64encode(contenu).decode("ascii")}}]
    if ext in ("xlsx", "xlsm"):
        texte = _texte_excel(contenu)
    elif ext == "docx":
        texte = _texte_docx(contenu)
    elif ext in ("csv", "txt"):
        texte = _texte_csv(contenu)
    elif ext == "xls":
        raise ExtractionError(f"{nom} : ancien format Excel (.xls). Enregistrez-le au format .xlsx puis réessayez.")
    else:
        raise ExtractionError(f"{nom} : format non pris en charge (PDF, image, Word .docx, Excel .xlsx ou CSV).")
    if not texte.strip():
        raise ExtractionError(f"{nom} : aucun texte lisible.")
    return [{"type": "text", "text": f"Document « {titre} » (contenu converti en texte) :\n{texte[:300000]}"}]


# ----------------------------------------------------------------------
# 1. RCCM et statuts
# ----------------------------------------------------------------------
_S = {"type": "string"}
_N = {"type": ["number", "null"]}
CHAMPS_ENTREPRISE = ["raison_sociale", "sigle", "forme_juridique", "date_creation", "rccm", "ncc", "siege",
                     "adresse", "ville", "telephone", "email", "domaine_activite", "activite", "duree_societe"]


def _objet(proprietes: Dict[str, Any]) -> Dict[str, Any]:
    return {"type": "object", "properties": proprietes, "required": list(proprietes), "additionalProperties": False}


SCHEMA_IDENTITE = _objet({
    "entreprise": _objet({**{c: _S for c in CHAMPS_ENTREPRISE}, "capital": _N,
                          "objet_social": {"type": "array", "items": _S}}),
    "dirigeants": {"type": "array", "items": _objet({
        "nom": _S, "fonction": _S, "qualite": _S, "nationalite": _S, "date_naissance": _S, "adresse": _S,
        "document": _S})},
    "associes": {"type": "array", "items": _objet({
        "nom": _S, "nature": _S, "nationalite": _S, "date_naissance": _S, "qualite": _S,
        "nombre_parts": _N, "montant": _N, "pourcentage": _N, "document": _S})},
    "valeur_nominale_part": _N,
    "provenance": {"type": "array", "items": _objet({"champ": _S, "document": _S})},
    "informations_absentes": {"type": "array", "items": _S},
    "divergences": {"type": "array", "items": _S},
    "remarques": _S,
})

CONSIGNE_IDENTITE = """Tu extrais les informations juridiques d'une entreprise ivoirienne (droit OHADA) à partir de ses
documents officiels : extrait ou déclaration d'immatriculation au RCCM et, s'ils sont fournis, statuts.

RÈGLE ABSOLUE : recopie uniquement ce qui est ÉCRIT dans les documents. N'invente, ne déduis et ne complète
AUCUNE information. Un champ absent reste une chaîne vide "" (ou null pour un nombre) et son nom est ajouté
à "informations_absentes" (libellé lisible, ex. « Numéro de compte contribuable (NCC) »).

"entreprise" :
- raison_sociale (dénomination sociale), sigle, forme_juridique (SARL, SA, SAS, SNC, SARLU...),
  date_creation (date de constitution ou, à défaut, d'immatriculation, au format JJ/MM/AAAA), rccm (numéro),
  ncc (numéro de compte contribuable / identifiant fiscal, seulement s'il est écrit), siege (siège social),
  adresse (adresse géographique ou postale complète), ville, telephone, email,
  capital (montant du capital social en FCFA, nombre entier), duree_societe (ex. « 99 ans »),
  objet_social (liste des activités telles qu'écrites, une par élément),
  activite (activité principale en une phrase fidèle au document), domaine_activite (domaine en quelques mots,
  repris du document).
"dirigeants" : gérant(s), président, directeur général, administrateurs... nom et prénoms, fonction, qualite
  (qualité juridique : gérant statutaire, PDG, administrateur, mandataire...), nationalité, date de naissance,
  adresse ; "document" = « RCCM » ou « Statuts » selon l'origine.
"associes" : chaque associé ou actionnaire figurant au capital (statuts en priorité) : nom, nature (personne
  physique / morale), nationalité, date de naissance, qualité, nombre_parts (parts ou actions), montant (apport
  en FCFA), pourcentage (part du capital en FRACTION entre 0 et 1, seulement s'il est écrit), "document".
  Un mandataire qui dépose la demande n'est pas un associé.
"valeur_nominale_part" : valeur nominale d'une part ou action si elle est écrite.
"provenance" : pour chaque champ de "entreprise" renseigné, le document d'où il vient (« RCCM » ou « Statuts »).
"divergences" : informations contradictoires entre RCCM et statuts (ex. capital différent), sinon liste vide.
"remarques" : difficulté de lecture éventuelle (page illisible, scan flou), sinon "".
"""


def _somme(vals: List[Optional[float]]) -> float:
    return sum(v for v in vals if v)


def controler_identite(data: Dict[str, Any]) -> Tuple[Dict[str, Any], List[str], List[str]]:
    """Controles deterministes et pourcentages calcules. Retourne (data, alertes, champs_calcules)."""
    alertes: List[str] = list(data.get("divergences") or [])
    calcules: List[str] = []
    e = data.get("entreprise") or {}
    capital = e.get("capital")
    associes = data.get("associes") or []
    vn = data.get("valeur_nominale_part")
    for a in associes:
        if a.get("pourcentage") is not None and a["pourcentage"] > 1.0001:
            a["pourcentage"] = a["pourcentage"] / 100  # ecrit en pourcentage entier
        if a.get("montant") in (None, 0) and a.get("nombre_parts") and vn:
            a["montant"] = a["nombre_parts"] * vn
            calcules.append(f"Apport de {a['nom']} calculé : {a['nombre_parts']:,.0f} parts × {vn:,.0f} FCFA."
                            .replace(",", " "))
    total_parts = _somme([a.get("nombre_parts") for a in associes])
    total_montants = _somme([a.get("montant") for a in associes])
    for a in associes:
        if a.get("pourcentage") is None:
            if total_parts and a.get("nombre_parts"):
                a["pourcentage"] = a["nombre_parts"] / total_parts
                calcules.append(f"Part de {a['nom']} calculée à partir du nombre de parts.")
            elif (capital or total_montants) and a.get("montant"):
                a["pourcentage"] = a["montant"] / (capital or total_montants)
                calcules.append(f"Part de {a['nom']} calculée à partir de son apport.")
    if associes:
        total_pct = _somme([a.get("pourcentage") for a in associes])
        if total_pct and abs(total_pct - 1) > 0.005:
            alertes.append(f"La répartition du capital totalise {total_pct * 100:.1f} % (et non 100 %) : "
                           "vérifiez la liste des associés.")
        if capital and total_montants and abs(total_montants - capital) > 1:
            alertes.append(f"La somme des apports ({total_montants:,.0f} FCFA) diffère du capital social "
                           f"({capital:,.0f} FCFA).".replace(",", " "))
    if capital is not None and capital <= 0:
        alertes.append("Capital social nul ou négatif : à vérifier.")
    if e.get("rccm") and not re.search(r"CI-[A-Z]{3}", e["rccm"].upper()):
        alertes.append(f"Numéro RCCM « {e['rccm']} » inhabituel (format attendu : CI-ABJ-...).")
    return data, alertes, calcules


def extraire_identite(fichiers: List[Tuple[str, bytes, str]]) -> Dict[str, Any]:
    """fichiers : liste de (nom, contenu, type) ou type = "rccm" ou "statuts"."""
    if not any(t == "rccm" for _, _, t in fichiers):
        raise ExtractionError("Le RCCM est obligatoire pour extraire l'identité de l'entreprise.")
    contenu: List[Dict[str, Any]] = []
    for nom, octets, typ in fichiers:
        contenu += blocs_document(nom, octets, "RCCM" if typ == "rccm" else "Statuts")
    contenu.append({"type": "text", "text": CONSIGNE_IDENTITE})
    data = appeler(contenu, schema=SCHEMA_IDENTITE, effort="medium")
    data, alertes, calcules = controler_identite(data)
    data["alertes"] = alertes
    data["calculs"] = calcules
    data["fichiers"] = [{"nom": n, "type": t} for n, _, t in fichiers]
    return data


# ----------------------------------------------------------------------
# 2. Compte d'exploitation et plan de tresorerie
# ----------------------------------------------------------------------
CATEGORIES = {
    "compte_exploitation": {
        "chiffre_affaires": "Chiffre d'affaires / ventes / prestations",
        "achats_matieres": "Achats de marchandises et de matières premières",
        "autres_achats": "Autres achats (fournitures, carburant, eau, électricité...)",
        "transport": "Transport",
        "services_exterieurs": "Services extérieurs (loyers, entretien, assurances, honoraires, télécoms...)",
        "impots_taxes": "Impôts et taxes",
        "salaires": "Salaires et charges sociales",
        "dotations": "Dotations aux amortissements",
        "frais_financiers": "Frais financiers / intérêts / échéances de crédit",
        "impot_benefice": "Impôt sur le bénéfice",
        "sous_total": "Ligne de sous-total ou de solde (marge, VA, EBE, résultat...)",
        "autre": "Autre ligne",
    },
    "plan_tresorerie": {
        "solde_initial": "Trésorerie / solde de début de période",
        "encaissement": "Encaissement (ventes, apports, emprunts...)",
        "decaissement": "Décaissement (achats, charges, salaires, échéances...)",
        "total_encaissements": "Total des encaissements",
        "total_decaissements": "Total des décaissements",
        "solde_final": "Trésorerie / solde de fin de période",
        "sous_total": "Autre sous-total",
        "autre": "Autre ligne",
    },
}
TITRES = {"compte_exploitation": "Compte d'exploitation prévisionnel", "plan_tresorerie": "Plan de trésorerie"}


def _schema_tableau(nature: str) -> Dict[str, Any]:
    return _objet({
        "titre": _S,
        "nature": {"type": "string", "enum": ["previsionnel", "historique"]},
        "periodes": {"type": "array", "items": _S},
        "lignes": {"type": "array", "items": _objet({
            "libelle": _S, "categorie": {"type": "string", "enum": list(CATEGORIES[nature])},
            "valeurs": {"type": "array", "items": _N}, "total": _N})},
        "remarques": _S,
    })


def _consigne_tableau(nature: str) -> str:
    cats = "\n".join(f"  - {k} : {v}" for k, v in CATEGORIES[nature].items())
    return f"""Tu lis un {TITRES[nature].lower()} d'une entreprise ivoirienne (montants en FCFA).
Recopie le tableau fidèlement, sans rien calculer, corriger ni compléter :
- "periodes" : les en-têtes de colonnes de période, dans l'ordre (ex. « Mois 1 »... ou « Janvier »... ou « 2026 »),
  SANS la colonne de total ;
- "lignes" : chaque ligne du tableau, dans l'ordre : libellé exact, catégorie, une valeur par période (null si la
  cellule est vide ou porte un tiret), et "total" = la colonne Total si elle existe (sinon null) ;
- catégories possibles :
{cats}
- montants : nombres entiers sans séparateur de milliers ; garde le signe imprimé ;
- "nature" : "previsionnel" si le document présente des prévisions, "historique" s'il s'agit de réalisations ;
- "titre" : le titre du tableau tel qu'écrit (ou une description courte) ;
- "remarques" : difficulté de lecture, colonnes ambiguës, sinon "".
N'invente aucune ligne ni aucune valeur absente du document."""


def controler_tableau(data: Dict[str, Any], nature: str) -> List[str]:
    """Incoherences detectees par le programme (totaux, soldes, valeurs manquantes)."""
    alertes: List[str] = []
    periodes = data.get("periodes") or []
    n = len(periodes)
    if not n:
        alertes.append("Aucune colonne de période reconnue.")
    for l in data.get("lignes", []):
        vals = l.get("valeurs") or []
        if n and len(vals) != n:
            alertes.append(f"« {l['libelle']} » : {len(vals)} valeur(s) pour {n} période(s).")
        manquantes = sum(1 for v in vals if v is None)
        if vals and manquantes == len(vals) and l.get("total") in (None, 0):
            alertes.append(f"« {l['libelle']} » : aucune valeur lue.")
        total = l.get("total")
        if total is not None and vals and l.get("categorie") not in ("solde_initial", "solde_final"):
            somme = sum(v or 0 for v in vals)
            if abs(somme - total) > max(1.0, 0.001 * abs(total)):
                alertes.append(f"« {l['libelle']} » : la somme des périodes ({somme:,.0f}) diffère du total "
                               f"imprimé ({total:,.0f}).".replace(",", " "))
    if nature == "plan_tresorerie" and n:
        lignes = data.get("lignes", [])

        def serie(cat: str) -> Optional[List[float]]:
            l = next((x for x in lignes if x.get("categorie") == cat), None)
            return [v or 0 for v in l["valeurs"]][:n] if l and l.get("valeurs") else None

        enc = serie("total_encaissements") or _somme_cat(lignes, "encaissement", n)
        dec = serie("total_decaissements") or _somme_cat(lignes, "decaissement", n)
        ini, fin = serie("solde_initial"), serie("solde_final")
        if ini and fin and enc and dec:
            for i in range(min(len(ini), len(fin), len(enc), len(dec))):
                attendu = ini[i] + enc[i] - dec[i]
                if abs(attendu - fin[i]) > max(1.0, 0.001 * abs(fin[i])):
                    alertes.append(f"{periodes[i]} : solde final {fin[i]:,.0f} ≠ solde initial + encaissements − "
                                   f"décaissements = {attendu:,.0f}.".replace(",", " "))
                    break
            for i in range(1, min(len(ini), len(fin))):
                if abs(ini[i] - fin[i - 1]) > 1:
                    alertes.append(f"{periodes[i]} : le solde initial ({ini[i]:,.0f}) ne reprend pas le solde final "
                                   f"de la période précédente ({fin[i - 1]:,.0f}).".replace(",", " "))
                    break
        if fin and any(v < 0 for v in fin):
            alertes.append("Trésorerie négative sur au moins une période : besoin de financement à court terme.")
    if nature == "compte_exploitation" and not any(l.get("categorie") == "chiffre_affaires"
                                                   for l in data.get("lignes", [])):
        alertes.append("Aucune ligne de chiffre d'affaires reconnue.")
    return alertes


def _somme_cat(lignes, cat: str, n: int) -> Optional[List[float]]:
    choisies = [l for l in lignes if l.get("categorie") == cat and l.get("valeurs")]
    if not choisies:
        return None
    return [sum((l["valeurs"][i] or 0) if i < len(l["valeurs"]) else 0 for l in choisies) for i in range(n)]


def _style(categorie: str) -> str:
    if categorie in ("total_encaissements", "total_decaissements", "sous_total"):
        return "sous_total"
    if categorie in ("solde_final",):
        return "total"
    return ""


def extraire_tableau(nom: str, contenu: bytes, nature: str) -> Dict[str, Any]:
    if nature not in CATEGORIES:
        raise ExtractionError(f"Type de tableau inconnu : {nature}.")
    data = appeler(blocs_document(nom, contenu, TITRES[nature]) + [{"type": "text", "text": _consigne_tableau(nature)}],
                   schema=_schema_tableau(nature), effort="medium")
    if not data.get("lignes"):
        raise IaError(f"Aucune ligne reconnue dans « {nom} » : vérifiez qu'il s'agit bien d'un "
                      f"{TITRES[nature].lower()}.")
    for l in data["lignes"]:
        l["style"] = _style(l.get("categorie", ""))
    data["alertes"] = controler_tableau(data, nature)
    data["fichier"] = nom
    data["titre"] = data.get("titre") or TITRES[nature]
    if nature == "compte_exploitation":
        data["proposition_projection"] = projection_depuis_compte(data)
    return data


def projection_depuis_compte(tableau: Dict[str, Any]) -> Dict[str, Any]:
    """Montants ANNUELS par rubrique tires d'un compte d'exploitation importe, pour pre-remplir la 1re annee des
    previsions (le consultant complete les annees suivantes). Les periodes infra-annuelles sont additionnees."""
    from .projection import RUBRIQUES_CHARGES

    produits, charges = [], []
    autres: Dict[str, float] = {}
    for l in tableau.get("lignes", []):
        cat = l.get("categorie")
        montant = l.get("total")
        if montant is None:
            montant = sum(v or 0 for v in l.get("valeurs") or [])
        if not montant:
            continue
        if cat == "chiffre_affaires":
            produits.append({"libelle": l["libelle"], "montant_annuel": round(montant)})
        elif cat in RUBRIQUES_CHARGES:
            charges.append({"libelle": l["libelle"], "rubrique": cat, "montant_annuel": round(montant)})
        elif cat in ("dotations", "frais_financiers", "impot_benefice"):
            autres[cat] = autres.get(cat, 0) + round(montant)
    return {"produits": produits, "charges": charges, "autres": autres,
            "periodes": len(tableau.get("periodes") or [])}
