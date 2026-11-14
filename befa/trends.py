"""
Analyse des exports CSV de Google Trends (Lot 3).

Le personnel BEFA exporte lui-meme les CSV depuis trends.google.com (bouton
"Telecharger" de chaque encadre) et les depose sur la page :
  - multiTimeline.csv    : evolution de l'interet dans le temps (obligatoire) ;
  - geoMap.csv           : interet par sous-region ;
  - relatedQueries.csv   : requetes associees (principales / en hausse) ;
  - relatedEntities.csv  : sujets associes.

Tout est calcule ici, sans IA : tendance, saisonnalite, pic, fiabilite,
regions. `faits()` produit des phrases chiffrees que l'IA (Lot 4) reformule
dans le Business Plan, et qui servent telles quelles si l'IA est indisponible.

Rappel methodologique (repris dans le PDF) : Google Trends donne un indice
RELATIF de 0 a 100 (100 = pic de popularite sur la periode et la zone
choisies), pas un volume de recherches ni une taille de marche.
"""
import csv
import io
import re
import statistics
from dataclasses import dataclass, field
from datetime import date, timedelta
from typing import Dict, List, Optional, Tuple

MOIS = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août",
        "septembre", "octobre", "novembre", "décembre"]

_ENTETES_TEMPS = {"semaine", "week", "mois", "month", "jour", "day", "time", "temps", "date", "heure", "hour"}
_ENTETES_GEO = {"région", "region", "sous-région", "subregion", "ville", "city", "pays", "country",
                "zone métropolitaine", "metro", "dma", "territoire"}
_SECTIONS_TOP = {"top", "principales", "principaux", "les plus populaires"}
_SECTIONS_HAUSSE = {"rising", "hausse", "en hausse", "en progression"}

SEUIL_TENDANCE = 0.15  # +/- 15 % entre la 1re et la derniere annee -> hausse / baisse
SEUIL_ZEROS = 0.10  # plus de 10 % de points a 0 -> faible volume, courbe bruitee


class TrendsError(ValueError):
    """Fichier Google Trends illisible ou de type inconnu (message pour le consultant)."""


# ----------------------------------------------------------------------
# Lecture
# ----------------------------------------------------------------------
def _decoder(contenu: bytes) -> str:
    encodages = ["utf-16"] if contenu[:2] in (b"\xff\xfe", b"\xfe\xff") else []
    for enc in encodages + ["utf-8-sig", "cp1252"]:
        try:
            texte = contenu.decode(enc)
            if "\x00" not in texte:
                return texte
        except UnicodeDecodeError:
            continue
    raise TrendsError("Encodage du fichier non reconnu : ré-exportez le CSV depuis Google Trends.")


def _lignes(texte: str) -> List[List[str]]:
    echantillon = texte[:2000]
    sep = ";" if echantillon.count(";") > echantillon.count(",") else ","
    return [[c.strip() for c in r] for r in csv.reader(io.StringIO(texte), delimiter=sep)]


def _valeur(cellule: str) -> Optional[float]:
    c = cellule.strip().replace(" ", "").replace("\xa0", "").replace(" ", "")
    if not c:
        return None
    if c.startswith("<"):  # "<1" : interet non nul mais inferieur a 1
        return 0.5
    c = c.rstrip("%").replace(",", ".")
    try:
        return float(c)
    except ValueError:
        return None


def _terme_et_zone(entete: str) -> Tuple[str, str]:
    m = re.match(r"^(.*?):\s*\((.*)\)\s*$", entete)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return entete.strip(), ""


def _date(cellule: str) -> Optional[date]:
    c = cellule.strip()
    for motif, fabrique in (
        (r"^(\d{4})-(\d{2})-(\d{2})", lambda m: date(int(m[1]), int(m[2]), int(m[3]))),
        (r"^(\d{4})-(\d{2})$", lambda m: date(int(m[1]), int(m[2]), 1)),
        (r"^(\d{2})/(\d{2})/(\d{4})", lambda m: date(int(m[3]), int(m[2]), int(m[1]))),
    ):
        m = re.match(motif, c)
        if m:
            try:
                return fabrique(m)
            except ValueError:
                return None
    return None


@dataclass
class SerieTemporelle:
    terme: str
    zone: str
    points: List[Tuple[date, float]]


@dataclass
class Evolution:
    categorie: str
    granularite: str  # "semaine", "mois", "jour"
    series: List[SerieTemporelle]

    @property
    def debut(self) -> date:
        return min(s.points[0][0] for s in self.series)

    @property
    def fin(self) -> date:
        return max(s.points[-1][0] for s in self.series)

    @property
    def zone(self) -> str:
        return next((s.zone for s in self.series if s.zone), "")


@dataclass
class InteretRegional:
    terme: str
    regions: List[Tuple[str, float]]  # trie par interet decroissant


@dataclass
class Associes:
    nature: str  # "requêtes" ou "sujets"
    principales: List[Tuple[str, str]]  # (libelle, valeur affichee)
    en_hausse: List[Tuple[str, str]]


@dataclass
class DonneesTrends:
    evolution: Optional[Evolution] = None
    regions: List[InteretRegional] = field(default_factory=list)
    associes: List[Associes] = field(default_factory=list)
    fichiers: List[str] = field(default_factory=list)


def _categorie(lignes: List[List[str]]) -> str:
    for r in lignes[:3]:
        if r and ":" in r[0] and not _date(r[0]):
            cle, _, val = r[0].partition(":")
            if cle.strip().lower() in ("catégorie", "category", "categorie"):
                return val.strip()
    return ""


def lire_fichier(contenu: bytes, nom: str = "") -> object:
    """Lit un export Google Trends et retourne Evolution, InteretRegional ou Associes."""
    lignes = _lignes(_decoder(contenu))
    if not lignes:
        raise TrendsError(f"{nom or 'Fichier'} : fichier vide.")
    categorie = _categorie(lignes)

    # Requetes / sujets associes : sections "PRINCIPALES" et "EN HAUSSE"
    sections = [i for i, r in enumerate(lignes) if len([c for c in r if c]) == 1
                and r[0].strip().lower() in _SECTIONS_TOP | _SECTIONS_HAUSSE]
    if sections:
        nature = "sujets" if "entit" in nom.lower() or "sujet" in nom.lower() or "topic" in nom.lower() else "requêtes"
        principales, hausse, courant = [], [], None
        for r in lignes:
            if len([c for c in r if c]) == 1 and r[0].strip().lower() in _SECTIONS_TOP:
                courant = principales
                continue
            if len([c for c in r if c]) == 1 and r[0].strip().lower() in _SECTIONS_HAUSSE:
                courant = hausse
                continue
            if courant is not None and len(r) >= 2 and r[0]:
                courant.append((r[0], r[1]))
        return Associes(nature, principales, hausse)

    # En-tete de tableau
    for i, r in enumerate(lignes):
        if len(r) < 2 or not r[0]:
            continue
        cle = r[0].strip().lower()
        if cle in _ENTETES_TEMPS or (i + 1 < len(lignes) and _date(lignes[i + 1][0] if lignes[i + 1] else "")):
            granularite = {"week": "semaine", "month": "mois", "day": "jour", "time": "", "temps": "", "date": ""}.get(cle, cle)
            series = []
            for j, entete in enumerate(r[1:], start=1):
                terme, zone = _terme_et_zone(entete)
                pts = []
                for ligne in lignes[i + 1:]:
                    if len(ligne) <= j:
                        continue
                    d, v = _date(ligne[0]), _valeur(ligne[j])
                    if d and v is not None:
                        pts.append((d, v))
                if pts:
                    series.append(SerieTemporelle(terme, zone, sorted(pts)))
            if not series:
                raise TrendsError(f"{nom or 'Fichier'} : aucune donnée datée trouvée.")
            if granularite not in ("semaine", "mois", "jour"):
                ecart = (series[0].points[1][0] - series[0].points[0][0]).days if len(series[0].points) > 1 else 7
                granularite = "jour" if ecart <= 1 else "semaine" if ecart <= 8 else "mois"
            return Evolution(categorie, granularite, series)
        if cle in _ENTETES_GEO:
            terme, _ = _terme_et_zone(r[1])
            regions = [(l[0], _valeur(l[1])) for l in lignes[i + 1:] if len(l) >= 2 and l[0]]
            regions = sorted(((n, v) for n, v in regions if v is not None), key=lambda x: -x[1])
            return InteretRegional(terme, regions)

    raise TrendsError(
        f"{nom or 'Fichier'} : format non reconnu. Exportez depuis Google Trends l'encadré « Évolution de "
        "l'intérêt pour la recherche » (multiTimeline.csv), « Intérêt par sous-région » (geoMap.csv) ou "
        "« Requêtes / Sujets associés »."
    )


def lire_fichiers(fichiers: List[Tuple[str, bytes]]) -> DonneesTrends:
    """Lit plusieurs exports ; les series de plusieurs multiTimeline sont fusionnees."""
    donnees = DonneesTrends()
    for nom, contenu in fichiers:
        obj = lire_fichier(contenu, nom)
        donnees.fichiers.append(nom)
        if isinstance(obj, Evolution):
            if donnees.evolution is None:
                donnees.evolution = obj
            else:
                donnees.evolution.series.extend(obj.series)
        elif isinstance(obj, InteretRegional):
            donnees.regions.append(obj)
        else:
            donnees.associes.append(obj)
    return donnees


# ----------------------------------------------------------------------
# Analyse
# ----------------------------------------------------------------------
def moyennes_mensuelles(serie: SerieTemporelle) -> List[Tuple[date, float]]:
    paquets: Dict[Tuple[int, int], List[float]] = {}
    for d, v in serie.points:
        paquets.setdefault((d.year, d.month), []).append(v)
    return [(date(a, m, 1), statistics.mean(v)) for (a, m), v in sorted(paquets.items())]


@dataclass
class AnalyseSerie:
    terme: str
    zone: str
    debut: date
    fin: date
    moyenne: float
    moyenne_premiere_annee: float
    moyenne_derniere_annee: float
    evolution: Optional[float]  # variation relative derniere / premiere annee
    tendance: str  # "en hausse", "en baisse", "stable"
    pente_annuelle: float  # points d'indice par an (regression lineaire sur les moyennes mensuelles)
    pic_date: date
    pic_valeur: float
    part_zeros: float
    fiable: bool
    saisonnalite: Optional[List[float]]  # 12 coefficients (1.0 = moyenne), None si < 2 ans de donnees
    mois_forts: List[int]  # numeros 1..12
    mois_faibles: List[int]
    amplitude_saisonniere: Optional[float]
    variation_recente: Optional[float]  # 3 derniers mois vs memes mois un an plus tot


def _regression(xs: List[float], ys: List[float]) -> float:
    mx, my = statistics.mean(xs), statistics.mean(ys)
    den = sum((x - mx) ** 2 for x in xs)
    return 0.0 if not den else sum((x - mx) * (y - my) for x, y in zip(xs, ys)) / den


def _moyenne_fenetre(points, debut: date, fin: date) -> Optional[float]:
    vals = [v for d, v in points if debut <= d < fin]
    return statistics.mean(vals) if vals else None


def analyser_serie(serie: SerieTemporelle) -> AnalyseSerie:
    pts = serie.points
    debut, fin = pts[0][0], pts[-1][0]
    valeurs = [v for _, v in pts]
    un_an = timedelta(days=365)

    premiere = _moyenne_fenetre(pts, debut, debut + un_an) or 0.0
    derniere = _moyenne_fenetre(pts, fin - un_an + timedelta(days=1), fin + timedelta(days=1)) or 0.0
    evolution = (derniere - premiere) / premiere if premiere else None
    if evolution is None or (fin - debut).days < 300:
        tendance = "stable"
    elif evolution >= SEUIL_TENDANCE:
        tendance = "en hausse"
    elif evolution <= -SEUIL_TENDANCE:
        tendance = "en baisse"
    else:
        tendance = "stable"

    mensuel = moyennes_mensuelles(serie)
    xs = [(d - mensuel[0][0]).days / 365.25 for d, _ in mensuel]
    pente = _regression(xs, [v for _, v in mensuel]) if len(mensuel) > 2 else 0.0

    pic_date, pic_valeur = max(pts, key=lambda p: p[1])
    part_zeros = sum(1 for v in valeurs if v == 0) / len(valeurs)
    moyenne = statistics.mean(valeurs)

    saison, forts, faibles, amplitude = None, [], [], None
    if (fin - debut).days >= 2 * 365 - 31 and moyenne > 0:
        par_mois: Dict[int, List[float]] = {}
        for d, v in mensuel:
            par_mois.setdefault(d.month, []).append(v)
        if len(par_mois) == 12:
            moy_mois = statistics.mean(statistics.mean(v) for v in par_mois.values())
            saison = [statistics.mean(par_mois[m]) / moy_mois if moy_mois else 1.0 for m in range(1, 13)]
            ordre = sorted(range(12), key=lambda i: saison[i])
            forts = [i + 1 for i in reversed(ordre[-3:])]
            faibles = [i + 1 for i in ordre[:3]]
            amplitude = max(saison) - min(saison)

    recent = None
    trois_mois = timedelta(days=91)
    r1 = _moyenne_fenetre(pts, fin - trois_mois, fin + timedelta(days=1))
    r0 = _moyenne_fenetre(pts, fin - un_an - trois_mois, fin - un_an + timedelta(days=1))
    if r1 is not None and r0:
        recent = (r1 - r0) / r0

    return AnalyseSerie(
        terme=serie.terme, zone=serie.zone, debut=debut, fin=fin, moyenne=moyenne,
        moyenne_premiere_annee=premiere, moyenne_derniere_annee=derniere, evolution=evolution,
        tendance=tendance, pente_annuelle=pente, pic_date=pic_date, pic_valeur=pic_valeur,
        part_zeros=part_zeros, fiable=part_zeros <= SEUIL_ZEROS, saisonnalite=saison,
        mois_forts=forts, mois_faibles=faibles, amplitude_saisonniere=amplitude, variation_recente=recent,
    )


# ----------------------------------------------------------------------
# Restitution
# ----------------------------------------------------------------------
def _pct(x: float) -> str:
    return f"{x * 100:+.0f} %".replace("+", "+").replace("-", "−")


def _date_fr(d: date) -> str:
    return f"{MOIS[d.month - 1]} {d.year}"


def _liste(mots: List[str]) -> str:
    return mots[0] if len(mots) == 1 else ", ".join(mots[:-1]) + " et " + mots[-1]


def source(donnees: DonneesTrends) -> str:
    ev = donnees.evolution
    if not ev:
        return "Source : Google Trends."
    zone = f", {ev.zone}" if ev.zone else ""
    return f"Source : Google Trends{zone}, {_date_fr(ev.debut)} – {_date_fr(ev.fin)}, données exportées par BEFA Consulting."


def faits(donnees: DonneesTrends) -> List[str]:
    """Phrases factuelles et chiffrees, sures (calculees), pour l'IA et pour le PDF."""
    out: List[str] = []
    ev = donnees.evolution
    if ev:
        analyses = [analyser_serie(s) for s in ev.series]
        zone = f" en {ev.zone}" if ev.zone else ""
        for a in analyses:
            phrase = (f"Sur la période {_date_fr(a.debut)} – {_date_fr(a.fin)}, l'intérêt de recherche pour "
                      f"« {a.terme} »{zone} est {a.tendance}")
            if a.evolution is not None:
                phrase += (f" : indice moyen de {a.moyenne_premiere_annee:.0f} sur la première année contre "
                           f"{a.moyenne_derniere_annee:.0f} sur la dernière ({_pct(a.evolution)})")
            out.append(phrase + ".")
            out.append(f"Le pic d'intérêt (indice 100 = maximum de la période) a été atteint la semaine du "
                       f"{a.pic_date.day} {_date_fr(a.pic_date)}." if ev.granularite == "semaine" else
                       f"Le pic d'intérêt a été atteint en {_date_fr(a.pic_date)}.")
            if a.saisonnalite and a.amplitude_saisonniere is not None:
                if a.amplitude_saisonniere < 0.25:
                    out.append(f"La recherche de « {a.terme} » est peu saisonnière (écart de "
                               f"{a.amplitude_saisonniere * 100:.0f} points entre le mois le plus fort et le plus faible).")
                else:
                    out.append(
                        f"Saisonnalité : les mois les plus actifs sont {_liste([MOIS[m - 1] for m in a.mois_forts])} "
                        f"(jusqu'à {_pct(max(a.saisonnalite) - 1)} par rapport à la moyenne), les plus calmes "
                        f"{_liste([MOIS[m - 1] for m in a.mois_faibles])} ({_pct(min(a.saisonnalite) - 1)})."
                    )
            if a.variation_recente is not None:
                out.append(f"Sur les trois derniers mois, l'intérêt varie de {_pct(a.variation_recente)} par rapport "
                           f"à la même période de l'année précédente.")
            if not a.fiable:
                out.append(f"Prudence : {a.part_zeros * 100:.0f} % des relevés valent 0, signe d'un faible volume de recherche "
                           f"sur « {a.terme} » ; la courbe est bruitée et seules les grandes tendances sont significatives.")
        if len(analyses) > 1:
            classement = sorted(analyses, key=lambda a: -a.moyenne)
            out.append("Popularité relative sur la période (indice moyen) : " +
                       ", ".join(f"« {a.terme} » {a.moyenne:.0f}" for a in classement) + ".")
    for r in donnees.regions:
        if r.regions:
            top = r.regions[:5]
            out.append(f"Régions où « {r.terme} » suscite le plus d'intérêt : " +
                       ", ".join(f"{n} ({v:.0f})" for n, v in top) + ".")
    for a in donnees.associes:
        if a.en_hausse:
            out.append(f"{a.nature.capitalize()} associées en forte progression : " +
                       ", ".join(f"« {n} » ({v})" for n, v in a.en_hausse[:5]) + ".")
        if a.principales:
            out.append(f"{a.nature.capitalize()} associées les plus fréquentes : " +
                       ", ".join(f"« {n} »" for n, _ in a.principales[:5]) + ".")
    return out


AVERTISSEMENT = ("Google Trends mesure un intérêt de recherche relatif (indice 0 à 100, 100 = pic de la période "
                 "et de la zone étudiées) ; il ne s'agit ni d'un volume de recherches ni d'une taille de marché. "
                 "Ces données éclairent les tendances et la saisonnalité de la demande et complètent l'étude de marché.")
