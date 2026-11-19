"""
Redaction des sections narratives du Business Plan BEFA (Lot 4).

Deux appels a Claude, puis des controles deterministes :

  1. RECHERCHE (outil de recherche web) : statistiques et faits du secteur.
     Seules les affirmations portant une CITATION (lien vers une page web
     reellement consultee par l'outil) sont retenues : une source inventee ne
     peut pas passer. Chaque fait recoit un identifiant [S1], [S2]...
  2. REDACTION (JSON impose) : sections au format BEFA, a partir du dossier,
     des faits sources, des constats Google Trends et des commentaires
     financiers (ecrits par le programme, cf. befa.commentaires).
  3. CONTROLES : tout nombre du texte introuvable dans les donnees fournies
     est signale au consultant ; les renvois [Sx] inconnus sont retires.
"""
import json
import re
from urllib.parse import unquote
from dataclasses import asdict, dataclass, field
from typing import Any, Dict, List, Optional, Tuple

from . import commentaires
from .dossier import DossierBEFA
from .ia import IaError, appeler
from .projection import Projection

# ----------------------------------------------------------------------
# Plan des sections : modele de reference "Dossier de credit O'IVOIRE",
# commun aux deux profils, consignes adaptees (StartUP / PME, ETI, GE).
# ----------------------------------------------------------------------
_PLAN = [
    # (cle, titre, consigne PME/ETI/GE, consigne STARTUP)
    ("historique", "Historique de la création",
     "Historique de l'entreprise à partir des informations du dossier (date et lieu de création, forme juridique, "
     "RCCM, capital, siège, évolutions marquantes). 1 à 2 paragraphes. N'ajoute aucune date absente du dossier.",
     "Genèse du projet : origine de l'idée, parcours et motivations du ou des promoteurs, stade d'avancement de la "
     "création (société constituée ou en cours de constitution). 1 à 2 paragraphes."),
    ("presentation_activite", "Présentation de l'activité",
     "Activités et prestations de l'entreprise, positionnement, zone d'intervention. 1 à 2 paragraphes ; puces pour "
     "la liste des activités.",
     "Activité envisagée, produits et services prévus, modèle économique, zone d'intervention. 1 à 2 paragraphes ; "
     "puces pour les produits et services."),
    ("clientele", "Clientèle",
     "Typologie de la clientèle et modes de règlement ; si une liste de clients chiffrée est fournie, commente le "
     "portefeuille et la dépendance commerciale en reprenant uniquement les chiffres fournis. 1 à 2 paragraphes.",
     "Clientèle cible (segments, localisation, besoins), canaux de vente envisagés, conditions de règlement prévues. "
     "1 à 2 paragraphes."),
    ("fournisseurs", "Fournisseurs",
     "Politique d'approvisionnement et principaux fournisseurs (nature des achats, conditions). 1 paragraphe.",
     "Sources d'approvisionnement envisagées (fournisseurs, équipementiers, matières premières). 1 paragraphe."),
    ("contexte_besoin", "Contexte",
     "Pourquoi l'entreprise sollicite ce financement maintenant : situation actuelle, contraintes, opportunité. "
     "2 à 4 paragraphes.",
     "Contexte et justification de la création : besoin du marché auquel répond le projet, opportunité, moyens "
     "nécessaires au démarrage. 2 à 4 paragraphes."),
    ("besoin_financement", "Besoin de financement",
     "Paragraphe sur la stratégie de croissance qui motive la demande, puis puces de l'utilisation des fonds "
     "sollicités (à partir des investissements et prêts fournis).",
     "Paragraphe sur les besoins de démarrage qui motivent la demande, puis puces de l'utilisation des fonds "
     "sollicités (à partir des investissements et prêts fournis)."),
    ("impact_besoin", "Impact du besoin sur les activités de l'entreprise",
     "Paragraphe d'introduction puis puces des impacts attendus (coûts, réactivité, qualité, capacité, marges).",
     "Paragraphe d'introduction puis puces des impacts attendus du financement sur le lancement et la montée en "
     "puissance de l'activité, l'emploi et la rentabilité."),
    ("contexte_mondial", "Contexte mondial du secteur",
     "Contexte mondial du secteur (tendances, enjeux, chiffres sourcés). 2 paragraphes, puces possibles.", None),
    ("contexte_africain", "Contexte africain du secteur",
     "Contexte africain / CEDEAO / UEMOA du secteur, avec chiffres sourcés. 2 paragraphes, puces possibles.", None),
    ("contexte_national", "Contexte et marché en Côte d'Ivoire",
     "Importance stratégique du secteur en Côte d'Ivoire, politiques et investissements publics, cadre "
     "réglementaire et incitations, chiffres sourcés. 3 à 4 paragraphes, puces possibles.", None),
    ("contexte_zone", "Contexte et analyse spécifique de la zone d'activité",
     "Situation de la zone d'activité (région, villes), facteurs structurels de la demande, opportunité économique "
     "locale. 2 à 3 paragraphes et puces.", None),
    ("analyse_demande", "Analyse du marché",
     "Analyse de la demande par segment de clientèle (résidentiel, entreprises, institutions... selon le secteur), "
     "tendances Google Trends si fournies (avec leur source et leur limite). 3 à 5 paragraphes, puces possibles.", None),
    ("analyse_prix", "Analyse des prix",
     "Déterminants des prix et prix pratiqués sur le marché (uniquement des chiffres fournis ou sourcés), "
     "positionnement prix de l'entreprise. 1 à 2 paragraphes et puces.", None),
    ("concurrence_directe", "Concurrence locale ou concurrence directe",
     "Caractéristiques et positionnement des concurrents directs (concurrents cités dans le dossier en priorité). "
     "1 paragraphe et puces.", None),
    ("concurrence_nationale", "Concurrence régionale et nationale",
     "Acteurs structurés au niveau régional et national, forces et limites. 1 paragraphe et puces.", None),
    ("opportunite_concurrentielle", "Opportunité concurrentielle",
     "Espace stratégique et avantages concurrentiels de l'entreprise. 1 paragraphe et puces.", None),
    ("conclusion_marche", "Conclusion et perspectives",
     "Synthèse de l'étude de marché et perspectives de croissance à moyen et long terme. 2 paragraphes.", None),
    ("conclusion", "Conclusion générale",
     "Synthèse : solidité de l'entreprise (en reprenant uniquement les commentaires financiers fournis), "
     "pertinence du besoin, capacité de remboursement, demande de financement. 2 à 3 paragraphes.",
     "Synthèse : pertinence du projet de création, potentiel du marché, viabilité prévisionnelle (en reprenant "
     "uniquement les commentaires financiers fournis, qui sont des PRÉVISIONS), demande de financement. "
     "2 à 3 paragraphes."),
]
SECTIONS = {
    "entreprise_existante": [(c, t, pme) for c, t, pme, _ in _PLAN],
    "nouveau_projet": [(c, t, st or pme) for c, t, pme, st in _PLAN],
}
# Cles des versions precedentes (plans COREAL / O'IVOIRE v1) : une redaction deja enregistree reste exploitable.
ANCIENNES_CLES = {
    "analyse_demande": ["marche"],
    "contexte_national": ["contexte_secteur", "introduction"],
    "contexte_besoin": ["justification"],
    "conclusion_marche": ["perspectives"],
    "besoin_financement": ["execution"],
}

# Variante "de base" de l'outil, volontairement : la variante 20260209 (filtrage
# dynamique par execution de code) ne rattache PAS de citations au texte final,
# or tout le controle des sources repose sur ces citations (verifie le 14/11/2026).
OUTIL_RECHERCHE = {
    "type": "web_search_20250305",
    "name": "web_search",
    "max_uses": 10,
    # sources ecartees a la demande de l'utilisateur (sous-domaines inclus : fr.wikipedia.org...)
    "blocked_domains": ["wikipedia.org"],
    # pas de "user_location" : le code pays CI n'est pas accepte par l'outil (erreur 400) ;
    # la consigne de recherche precise la Cote d'Ivoire.
}

SOURCES_FIABLES = (
    "Institut National de la Statistique (INS) de Côte d'Ivoire, ministères ivoiriens, CEPICI, BCEAO, UEMOA, "
    "CEDEAO, Banque mondiale, FMI, BAD, FAO, OCDE, CNUCED, interprofessions et organismes de filière (ex. FIRCA, "
    "AIPH, Conseil Café-Cacao), agences de presse et presse économique reconnue, rapports d'entreprises cotées"
)


# ----------------------------------------------------------------------
# 1. Recherche web sourcee
# ----------------------------------------------------------------------
@dataclass
class Source:
    id: str
    titre: str
    url: str


@dataclass
class Fait:
    texte: str
    sources: List[str]  # identifiants S1, S2...
    extraits: List[str] = field(default_factory=list)  # passages cites des pages sources (verification)


@dataclass
class Recherche:
    faits: List[Fait] = field(default_factory=list)
    sources: List[Source] = field(default_factory=list)
    consultees: List[Tuple[str, str]] = field(default_factory=list)  # (titre, url) de toutes les pages trouvees

    def source(self, sid: str) -> Optional[Source]:
        return next((s for s in self.sources if s.id == sid), None)


def _consigne_recherche(d: DossierBEFA) -> str:
    from .secteurs import reperes

    p, e = d.projet, d.entreprise
    sujets = {
        "nouveau_projet": [
            "le contexte mondial et africain du secteur (production, consommation, échanges, prix) ;",
            "la situation du secteur en Côte d'Ivoire (production, principaux acteurs, importations, plans "
            "et politiques publiques) ;",
            "le cadre réglementaire et les incitations à l'investissement applicables (Code des investissements, "
            "CEPICI, normes environnementales et sanitaires) ;",
            f"la zone d'implantation ({p.localisation or 'à préciser'}) : ressources, population, infrastructures ;",
            "la demande par segment, les prix pratiqués et la concurrence sur le marché visé.",
        ],
        "entreprise_existante": [
            "le contexte mondial et africain du secteur ;",
            "le marché ivoirien du secteur (taille, croissance, programmes publics, acteurs) ;",
            f"la zone d'activité ({p.localisation or e.ville or 'à préciser'}) : besoins, démographie, projets ;",
            "la demande par segment, les prix pratiqués et la concurrence.",
        ],
    }[d.variante]
    axes = reperes(p.secteur_code, p.secteur)
    return (
        f"Tu prépares l'étude de marché d'un Business Plan rédigé par le cabinet BEFA Consulting (Abidjan) pour "
        f"{'le projet de création' if d.startup else 'l’entreprise'} « {d.raison_sociale} ».\n"
        f"Projet : {p.intitule}. Secteur : {p.secteur}{' — ' + p.sous_secteur if p.sous_secteur else ''}. "
        f"Localisation : {p.localisation or e.ville}. "
        f"Marché visé : {p.zone_marche}. Produits / services : "
        f"{', '.join(x.nom for x in d.produits) or p.objet or 'voir secteur'}.\n\n"
        "Fais des recherches sur le web pour documenter :\n- " + "\n- ".join(sujets) + "\n\n" +
        (axes + "\n\n" if axes else "") +
        f"Privilégie les sources fiables et récentes (moins de 5 ans si possible) : {SOURCES_FIABLES}. "
        "Évite les blogs, forums et sites sans auteur identifiable.\n\n"
        "Restitue tes constats en français, sous forme de courts paragraphes indépendants : chacun énonce un fait "
        "précis (chiffre, année de la donnée, périmètre géographique) et s'appuie sur la page qui le contient. "
        "N'écris aucun chiffre que tu n'as pas trouvé dans une source. Si une information est introuvable, "
        "dis-le simplement, sans l'estimer."
    )


def _attr(obj, nom, defaut=None):
    if isinstance(obj, dict):
        return obj.get(nom, defaut)
    return getattr(obj, nom, defaut)


def _titre_lisible(titre: str) -> str:
    """Certains titres arrivent encodes comme une URL ("G%C3%B4h") : on les decode."""
    titre = titre.strip()
    if re.search(r"%[0-9A-Fa-f]{2}", titre):
        titre = unquote(titre).replace("_", " ")
    return titre


def analyser_reponse_recherche(blocs: List[Any]) -> Recherche:
    """Extrait les faits cites et les sources d'une reponse avec recherche web."""
    res = Recherche()
    par_url: Dict[str, Source] = {}

    for b in blocs:
        if _attr(b, "type") == "web_search_tool_result":
            contenu = _attr(b, "content")
            if isinstance(contenu, list):  # une erreur de recherche est un objet, pas une liste
                for r in contenu:
                    url = _attr(r, "url")
                    if url:
                        res.consultees.append((_attr(r, "title", "") or url, url))

    # Les citations portent sur des fragments de phrase : on regroupe les blocs
    # de texte jusqu'a la fin de la phrase pour garder le contexte du fait.
    tampon, cits = "", []

    def vider():
        nonlocal tampon, cits
        texte = re.sub(r"\s+", " ", tampon).strip()
        citations = [c for c in cits if _attr(c, "url")]
        tampon, cits = "", []
        if not citations or len(texte) < 25:
            return  # affirmation non sourcee : ecartee
        ids, extraits = [], []
        for c in citations:
            url = _attr(c, "url")
            if url not in par_url:
                par_url[url] = Source(f"S{len(par_url) + 1}", _titre_lisible(_attr(c, "title") or url), url)
                res.sources.append(par_url[url])
            if par_url[url].id not in ids:
                ids.append(par_url[url].id)
            extrait = re.sub(r"\s+", " ", _attr(c, "cited_text", "") or "").strip()
            if extrait and extrait not in extraits:
                extraits.append(extrait)
        if not any(f.texte == texte for f in res.faits):
            res.faits.append(Fait(texte, ids, extraits))

    for b in blocs:
        if _attr(b, "type") != "text":
            continue
        morceau = _attr(b, "text", "") or ""
        tampon += morceau
        cits += list(_attr(b, "citations") or [])
        fin = morceau.rstrip()
        if "\n" in morceau or fin.endswith((".", "!", "?", ":", ";")):
            vider()
    vider()
    return res


def rechercher(d: DossierBEFA) -> Recherche:
    blocs = appeler([{"type": "text", "text": _consigne_recherche(d)}], outils=[OUTIL_RECHERCHE],
                    effort="medium", max_tokens=32000)
    return analyser_reponse_recherche(blocs)


# ----------------------------------------------------------------------
# 2. Commentaires financiers (programme) et redaction (IA)
# ----------------------------------------------------------------------
def commentaires_financiers(d: DossierBEFA, proj: Optional[Projection]) -> Dict[str, List[str]]:
    nom = d.entreprise.sigle or d.entreprise.raison_sociale
    out: Dict[str, List[str]] = {}
    if proj is not None:
        out["financement"] = commentaires.financement(proj, nom)
        out["compte_exploitation"] = commentaires.compte_exploitation(proj, nom)
        out["plan_financement"] = commentaires.plan_financement(proj, nom)
        out["couverture_dette"] = commentaires.couverture_dette(proj)
    h = d.analyse_historique
    if h is not None:
        out["analyse_bilan"] = commentaires.analyse_bilan(h, nom)
        out["analyse_compte_resultat"] = commentaires.analyse_compte_resultat(h, nom)
    return out


def _schema(variante: str) -> dict:
    cles = [c for c, _, _ in SECTIONS[variante]]
    liste = {"type": "array", "items": {"type": "string"}}
    return {
        "type": "object",
        "properties": {
            "sections": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {"cle": {"type": "string", "enum": cles}, "paragraphes": liste, "puces": liste},
                    "required": ["cle", "paragraphes", "puces"],
                    "additionalProperties": False,
                },
            },
            "swot": {
                "type": "object",
                "properties": {"forces": liste, "faiblesses": liste, "opportunites": liste, "menaces": liste},
                "required": ["forces", "faiblesses", "opportunites", "menaces"],
                "additionalProperties": False,
            },
        },
        "required": ["sections", "swot"],
        "additionalProperties": False,
    }


SYSTEME_REDACTION = """Tu es consultant senior au cabinet BEFA Consulting (Bureau d'Études Financières et d'Assistance, Abidjan), fondé par d'anciens cadres de banque. Tu rédiges les parties narratives d'un Business Plan destiné à un comité de crédit bancaire en Côte d'Ivoire.

Style : français soutenu, factuel et convaincant, phrases complètes, vocabulaire bancaire et OHADA ; pas de marketing creux, pas de superlatifs gratuits. Montants en FCFA avec espaces comme séparateurs de milliers.

Profil du demandeur :
- STARTUP (entreprise en création) : aucun historique comptable ; présente le projet de création, le promoteur et les PRÉVISIONS ; ne présente jamais une estimation comme une donnée réalisée ; le bilan d'ouverture est prévisionnel.
- PME, ETI, GE (entreprise existante) : appuie-toi sur l'historique, les performances passées et le projet de développement.

Règles impératives sur les chiffres :
- N'utilise QUE des chiffres présents dans les données fournies : dossier client, faits sourcés, constats Google Trends, commentaires financiers. N'invente, n'estime ni ne recalcule aucun chiffre.
- Après toute phrase qui reprend un fait sourcé, ajoute son renvoi entre crochets, ex. [S3] ou [S2][S5]. N'utilise que les identifiants fournis.
- Les commentaires financiers sont déjà rédigés et figureront à côté des tableaux : n'en reprends l'essentiel que dans la conclusion.
- Si une information manque, rédige sans elle (ou de façon qualitative) ; ne la comble pas par une supposition chiffrée.
- N'invente aucune donnée administrative (date, numéro RCCM, capital, nom de dirigeant ou d'associé) absente du dossier."""


def _contexte(d: DossierBEFA, rech: Recherche, coms: Dict[str, List[str]]) -> str:
    faits = "\n".join("".join(f"[{s}]" for s in f.sources) + f" {f.texte}" for f in rech.faits) \
        or "(aucun fait sourcé disponible)"
    trends = "\n".join(f"- {x}" for x in d.trends_faits) or "(pas de données Google Trends)"
    if d.trends_faits and d.trends_source:
        trends += f"\n{d.trends_source}"
    fin = "\n".join(f"{cle} :\n" + "\n".join(f"- {p}" for p in paras) for cle, paras in coms.items()) or "(aucun)"
    return (
        "DOSSIER CLIENT (JSON) :\n" + json.dumps(d.resume_pour_ia(), ensure_ascii=False, indent=1) +
        "\n\nFAITS SOURCÉS (identifiants entre crochets) :\n" + faits +
        "\n\nCONSTATS GOOGLE TRENDS (calculés par le programme) :\n" + trends +
        "\n\nCOMMENTAIRES FINANCIERS (rédigés par le programme) :\n" + fin
    )


def _consigne_redaction(d: DossierBEFA) -> str:
    from .secteurs import reperes

    plan = "\n".join(f"- {cle} — « {titre} » : {consigne}" for cle, titre, consigne in SECTIONS[d.variante])
    axes = reperes(d.projet.secteur_code, d.projet.secteur)
    profil = ("PROFIL : STARTUP (entreprise en création, Business Plan orienté création et prévisions)."
              if d.startup else "PROFIL : PME, ETI, GE (entreprise existante, Business Plan intégrant son historique, "
                                "ses performances et son projet de développement).")
    return (
        profil + "\n" + (axes + " Adapte l'analyse du marché, de la concurrence, des risques et les recommandations "
                         "à ce secteur.\n" if axes else "") +
        "\nRédige les sections suivantes du Business Plan (une entrée par clé, dans cet ordre). Mets le texte "
        "courant dans \"paragraphes\" et les énumérations dans \"puces\" (liste vide si aucune) :\n" + plan +
        "\n\nRemplis aussi \"swot\" : 3 à 5 éléments par rubrique (forces, faiblesses, opportunités, menaces), "
        "propres à ce projet, chacun au format « Élément identifié : impact sur le projet » "
        "(ex. « Demande structurelle élevée : volume d'activité soutenu et marché durable »)."
    )


# ----------------------------------------------------------------------
# 3. Controles deterministes
# ----------------------------------------------------------------------
_UNITES = {"milliard": 1e9, "milliards": 1e9, "million": 1e6, "millions": 1e6, "mille": 1e3,
           "mds": 1e9, "md": 1e9, "m": 1e6}
_NOMBRE = re.compile(
    r"(?<![\w])(\d{1,3}(?:[   .]\d{3})+(?:,\d+)?|\d+(?:[.,]\d+)?)"
    r"(?:\s?(milliards?|millions?|mille|mds|md)\b)?",
    re.I,
)


def _nombres(texte: str) -> List[Tuple[float, str]]:
    out = []
    for m in _NOMBRE.finditer(texte):
        brut, unite = m.group(1), (m.group(2) or "").lower()
        if re.fullmatch(r"\d{1,3}(?:[   .]\d{3})+(?:,\d+)?", brut):
            val = float(re.sub(r"[   .]", "", brut).replace(",", "."))
        else:
            val = float(brut.replace(",", "."))
        out.append((val * _UNITES.get(unite, 1), m.group(0)))
    return out


def _anodin(val: float, extrait: str) -> bool:
    if val.is_integer() and 1900 <= val <= 2100 and len(extrait.strip()) == 4:
        return True  # annee
    return val.is_integer() and val <= 10  # petits entiers (enumerations, durees courtes)


def verifier_chiffres(texte: str, corpus: str) -> List[str]:
    """Nombres du texte introuvables (a 0,6 % pres, au signe pres) dans le corpus des donnees fournies."""
    texte = re.sub(r"\[S\d+\]", "", texte)
    reference = [abs(v) for v, _ in _nombres(corpus)]
    inconnus = []
    for val, extrait in _nombres(texte):
        if _anodin(val, extrait):
            continue
        a = abs(val)
        if any(abs(a - r) <= 0.006 * max(r, 1) for r in reference):
            continue
        if extrait.strip() not in inconnus:
            inconnus.append(extrait.strip())
    return inconnus


@dataclass
class SectionRedigee:
    cle: str
    titre: str
    paragraphes: List[str]
    puces: List[str]


@dataclass
class Redaction:
    variante: str
    sections: List[SectionRedigee]
    swot: Dict[str, List[str]]
    commentaires_financiers: Dict[str, List[str]]
    sources: List[Source]  # sources effectivement citees, dans l'ordre de numerotation
    faits: List[Fait]
    consultees: List[Tuple[str, str]]
    alertes: List[str]

    def section(self, cle: str) -> Optional[SectionRedigee]:
        for c in [cle] + ANCIENNES_CLES.get(cle, []):
            s = next((s for s in self.sections if s.cle == c and (s.paragraphes or s.puces)), None)
            if s:
                return s
        return None

    def vers_json(self) -> dict:
        return asdict(self)


def _corpus(d: DossierBEFA, rech: Recherche, coms: Dict[str, List[str]], proj: Optional[Projection]) -> str:
    morceaux = [json.dumps(d.resume_pour_ia(), ensure_ascii=False)]
    morceaux += [f.texte for f in rech.faits]
    morceaux += d.trends_faits
    morceaux += [p for paras in coms.values() for p in paras]
    if proj is not None:
        ce = proj.compte
        for serie in (ce.chiffre_affaires, ce.valeur_ajoutee, ce.ebe, ce.resultat_net, ce.caf, ce.dotations,
                      proj.plan.cumul, *ce.charges_par_rubrique.values()):
            morceaux.append(" ".join(str(v) for v in serie))
        morceaux.append(" ".join(str(i.montant) for i in proj.entree.investissements))
        morceaux.append(str(proj.cout_total_investissements))
    return "\n".join(morceaux)


def assembler(d: DossierBEFA, brut: dict, rech: Recherche, coms: Dict[str, List[str]],
              proj: Optional[Projection]) -> Redaction:
    """Met en forme la reponse IA et applique les controles (deterministe, testable sans reseau)."""
    titres = {cle: titre for cle, titre, _ in SECTIONS[d.variante]}
    alertes: List[str] = []
    ids_valides = {s.id for s in rech.sources}
    cites: List[str] = []

    def nettoyer(t: str, ou: str) -> str:
        for sid in re.findall(r"\[(S\d+)\]", t):
            if sid not in ids_valides:
                alertes.append(f"{ou} : renvoi [{sid}] inconnu supprimé.")
            elif sid not in cites:
                cites.append(sid)
        return re.sub(r"\[(S\d+)\]", lambda m: m.group(0) if m.group(1) in ids_valides else "", t).strip()

    par_cle = {s.get("cle"): s for s in brut.get("sections", [])}
    sections = []
    for cle, titre in titres.items():
        s = par_cle.get(cle) or next((par_cle[x] for x in ANCIENNES_CLES.get(cle, []) if x in par_cle), None)
        if not s:
            alertes.append(f"Section « {titre} » non rédigée par l'IA : à compléter manuellement.")
            sections.append(SectionRedigee(cle, titre, [], []))
            continue
        sections.append(SectionRedigee(
            cle, titre,
            [nettoyer(p, titre) for p in s.get("paragraphes", []) if p.strip()],
            [nettoyer(p, titre) for p in s.get("puces", []) if p.strip()],
        ))
    swot = {k: [nettoyer(x, "SWOT") for x in v if x.strip()] for k, v in brut.get("swot", {}).items()}

    corpus = _corpus(d, rech, coms, proj)
    for s in sections:
        for bloc in s.paragraphes + s.puces:
            for nombre in verifier_chiffres(bloc, corpus):
                alertes.append(f"{s.titre} : le chiffre « {nombre} » ne figure dans aucune donnée fournie — "
                               f"à vérifier ou supprimer.")
    for k, v in swot.items():
        for x in v:
            for nombre in verifier_chiffres(x, corpus):
                alertes.append(f"SWOT : le chiffre « {nombre} » ne figure dans aucune donnée fournie — à vérifier.")

    sources = [s for s in rech.sources if s.id in cites]
    return Redaction(d.variante, sections, swot, coms, sources, rech.faits, rech.consultees, alertes)


def rediger(d: DossierBEFA, proj: Optional[Projection] = None, rech: Optional[Recherche] = None,
            recherche_web: bool = True) -> Redaction:
    """Chaine complete : recherche web (facultative), redaction, controles. Leve IaError si la redaction echoue ;
    un echec de la seule recherche web n'empeche pas la redaction (alerte ajoutee)."""
    alertes_recherche = []
    if rech is None:
        rech = Recherche()
        if recherche_web:
            try:
                rech = rechercher(d)
            except IaError as e:
                alertes_recherche.append(f"Recherche web indisponible ({e}) : texte rédigé sans statistiques sourcées.")
    coms = commentaires_financiers(d, proj)
    brut = appeler(
        [{"type": "text", "text": _contexte(d, rech, coms)}, {"type": "text", "text": _consigne_redaction(d)}],
        systeme=SYSTEME_REDACTION, schema=_schema(d.variante), effort="high", max_tokens=64000,
    )
    red = assembler(d, brut, rech, coms, proj)
    red.alertes[0:0] = alertes_recherche
    return red


def redaction_depuis_json(d: dict) -> Redaction:
    """Relit une redaction (eventuellement corrigee par le consultant) au format de Redaction.vers_json()."""
    return Redaction(
        variante=d["variante"],
        sections=[SectionRedigee(**s) for s in d.get("sections", [])],
        swot={k: list(v) for k, v in d.get("swot", {}).items()},
        commentaires_financiers={k: list(v) for k, v in d.get("commentaires_financiers", {}).items()},
        sources=[Source(**s) for s in d.get("sources", [])],
        faits=[Fait(**f) for f in d.get("faits", [])],
        consultees=[tuple(x) for x in d.get("consultees", [])],
        alertes=list(d.get("alertes", [])),
    )
