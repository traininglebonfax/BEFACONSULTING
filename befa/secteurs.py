"""
Referentiel des secteurs d'activite (liste deroulante de la page "Projet").

Chaque secteur porte des sous-secteurs et des REPERES d'analyse transmis a
l'IA (recherche documentaire et redaction) : indicateurs a documenter et
risques typiques. Ce sont des axes de travail, jamais des chiffres : l'IA
ne doit citer que des chiffres trouves dans une source.
"""
from typing import Dict, List, Optional

SECTEURS: List[Dict] = [
    {"code": "agriculture", "libelle": "Agriculture, élevage et pêche",
     "sous_secteurs": ["Cultures vivrières", "Cultures de rente (cacao, café, anacarde, hévéa, palmier)",
                       "Maraîchage", "Élevage avicole", "Élevage bovin, ovin, porcin", "Pêche et aquaculture",
                       "Production de semences et intrants"],
     "indicateurs": "production nationale et rendements, prix bord-champ et prix de marché, importations, "
                    "programmes publics agricoles (PNIA), accès aux intrants et à l'irrigation",
     "risques": "aléas climatiques, maladies et épizooties, volatilité des prix, saisonnalité des revenus, "
                "pertes post-récolte"},
    {"code": "agro_industrie", "libelle": "Agro-industrie et transformation alimentaire",
     "sous_secteurs": ["Huilerie", "Transformation du cacao", "Transformation de l'anacarde", "Boissons et jus",
                       "Minoterie et boulangerie industrielle", "Produits laitiers", "Conserverie",
                       "Aliments pour bétail"],
     "indicateurs": "taux de transformation locale, importations de produits finis, capacité installée, "
                    "disponibilité des matières premières, normes sanitaires (CODINORM)",
     "risques": "approvisionnement en matières premières, saisonnalité, qualité et normes, coût de l'énergie, "
                "concurrence des importations"},
    {"code": "industrie", "libelle": "Industrie manufacturière",
     "sous_secteurs": ["Plasturgie", "Métallurgie et transformation des métaux", "Chimie et cosmétiques",
                       "Matériaux de construction", "Emballages", "Bois et ameublement", "Imprimerie"],
     "indicateurs": "indice de la production industrielle, importations concurrentes, coût des facteurs, "
                    "zones industrielles, incitations du Code des investissements",
     "risques": "coût de l'énergie et des intrants importés, change, maintenance des équipements, "
                "concurrence asiatique"},
    {"code": "mines", "libelle": "Industrie extractive et mines",
     "sous_secteurs": ["Or", "Manganèse", "Carrières et granulats", "Pétrole et gaz", "Services miniers"],
     "indicateurs": "production minière, permis d'exploitation, Code minier, cours mondiaux, sous-traitance locale",
     "risques": "volatilité des cours, réglementation et permis, risques environnementaux et sociaux, "
                "sécurité des sites"},
    {"code": "btp", "libelle": "BTP, construction et travaux publics",
     "sous_secteurs": ["Bâtiment", "Travaux publics et routes", "Voirie et réseaux divers (VRD)",
                       "Forage et hydraulique", "Génie civil", "Second œuvre"],
     "indicateurs": "investissements publics (PND), carnet de marchés publics, prix du ciment et du fer, "
                    "déficit de logements",
     "risques": "retards de paiement de l'État, dépendance aux marchés publics, hausse du coût des matériaux, "
                "pénalités de retard"},
    {"code": "immobilier", "libelle": "Immobilier et gestion foncière",
     "sous_secteurs": ["Promotion immobilière", "Gestion locative", "Aménagement foncier", "Agence immobilière"],
     "indicateurs": "déficit de logements, prix du foncier, loyers, accès au crédit immobilier, sécurisation foncière",
     "risques": "litiges fonciers, délais d'obtention des ACD, solvabilité des acquéreurs, taux d'occupation"},
    {"code": "transport", "libelle": "Transport et logistique",
     "sous_secteurs": ["Transport de marchandises", "Transport de personnes", "Logistique et entreposage",
                       "Transit et commission en douane", "Transport maritime et portuaire", "Livraison urbaine"],
     "indicateurs": "trafic portuaire (Abidjan, San Pedro), parc de véhicules, corridors régionaux, prix du carburant",
     "risques": "prix du carburant, état des routes, sinistralité, vieillissement du parc, tracasseries routières"},
    {"code": "commerce", "libelle": "Commerce général et distribution",
     "sous_secteurs": ["Commerce de gros", "Commerce de détail", "Grande distribution", "Import-export",
                       "Distribution de produits pétroliers", "Quincaillerie"],
     "indicateurs": "consommation des ménages, importations, réseau de distribution, marges commerciales",
     "risques": "rotation des stocks, crédit clients, concurrence informelle, change sur importations"},
    {"code": "tic", "libelle": "Technologies de l'information et communication",
     "sous_secteurs": ["Développement logiciel", "Intégration de systèmes", "Infogérance et cloud",
                       "Cybersécurité", "Fourniture de matériel informatique", "Formation numérique"],
     "indicateurs": "taux de pénétration d'internet, transformation numérique de l'administration, "
                    "dépenses IT des entreprises, ARTCI",
     "risques": "obsolescence technologique, dépendance à quelques gros clients, rareté des compétences, "
                "cybersécurité"},
    {"code": "telecom", "libelle": "Télécommunications",
     "sous_secteurs": ["Opérateur", "Infrastructures et pylônes", "Fibre optique", "Services à valeur ajoutée"],
     "indicateurs": "abonnés mobiles et internet (ARTCI), couverture réseau, investissements des opérateurs",
     "risques": "réglementation et licences, intensité capitalistique, concurrence des grands opérateurs"},
    {"code": "fintech", "libelle": "FinTech et services financiers",
     "sous_secteurs": ["Paiement mobile", "Microfinance", "Agrégation de paiements", "Crédit digital",
                       "Transfert d'argent"],
     "indicateurs": "comptes de monnaie électronique (BCEAO), taux de bancarisation, réglementation BCEAO "
                    "sur les établissements de paiement",
     "risques": "agrément BCEAO, fraude, conformité LBC/FT, dépendance aux opérateurs télécoms"},
    {"code": "banque_assurance", "libelle": "Banque et assurance",
     "sous_secteurs": ["Banque", "Assurance", "Courtage", "Établissement financier"],
     "indicateurs": "taux de bancarisation, primes d'assurance (CIMA), crédit à l'économie (BCEAO)",
     "risques": "réglementation prudentielle, risque de crédit, conformité"},
    {"code": "energie", "libelle": "Énergie et énergies renouvelables",
     "sous_secteurs": ["Solaire photovoltaïque", "Biomasse", "Hydroélectricité", "Distribution d'énergie",
                       "Efficacité énergétique", "Gaz domestique"],
     "indicateurs": "taux d'électrification, mix énergétique et objectifs nationaux d'énergies renouvelables, "
                    "tarifs CIE, programmes d'électrification rurale",
     "risques": "cadre tarifaire, change sur équipements importés, maintenance, solvabilité des clients"},
    {"code": "eau_environnement", "libelle": "Eau, assainissement et environnement",
     "sous_secteurs": ["Forage et adduction d'eau", "Traitement et embouteillage d'eau",
                       "Assainissement", "Études environnementales"],
     "indicateurs": "taux d'accès à l'eau potable, programmes publics (hydraulique villageoise), "
                    "investissements du secteur",
     "risques": "conditions hydrogéologiques, dépendance aux marchés publics, coûts d'équipement"},
    {"code": "sante", "libelle": "Santé et industrie pharmaceutique",
     "sous_secteurs": ["Clinique et centre de santé", "Pharmacie", "Laboratoire d'analyses",
                       "Production pharmaceutique", "Distribution de matériel médical"],
     "indicateurs": "couverture maladie universelle (CMU), densité médicale, importations de médicaments, "
                    "autorité de régulation (AIRP)",
     "risques": "autorisations et normes, recrutement de praticiens, coût des équipements, tiers payant"},
    {"code": "education", "libelle": "Éducation et formation professionnelle",
     "sous_secteurs": ["Enseignement préscolaire et primaire", "Enseignement secondaire",
                       "Enseignement supérieur", "Formation professionnelle", "Certifications"],
     "indicateurs": "effectifs scolarisés, taux de scolarisation, demande de compétences, agréments du ministère",
     "risques": "agréments, recouvrement des frais de scolarité, qualité du corps enseignant, subventions"},
    {"code": "tourisme", "libelle": "Tourisme, hôtellerie et restauration",
     "sous_secteurs": ["Hôtellerie", "Restauration", "Agence de voyages", "Événementiel", "Loisirs"],
     "indicateurs": "arrivées touristiques, taux d'occupation hôtelière, stratégie Sublime Côte d'Ivoire",
     "risques": "saisonnalité, sécurité, concurrence, normes d'hygiène"},
    {"code": "services_entreprises", "libelle": "Services aux entreprises et conseil",
     "sous_secteurs": ["Conseil et audit", "Comptabilité", "Ressources humaines et intérim", "Nettoyage",
                       "Sécurité privée", "Marketing et communication"],
     "indicateurs": "dynamisme des créations d'entreprises (CEPICI), externalisation, commande publique",
     "risques": "dépendance à quelques clients, délais de paiement, rétention des talents"},
    {"code": "artisanat_culture", "libelle": "Artisanat et économie culturelle",
     "sous_secteurs": ["Artisanat d'art", "Industries culturelles et créatives", "Audiovisuel", "Musique",
                       "Édition"],
     "indicateurs": "poids de l'artisanat, politiques culturelles, débouchés à l'export",
     "risques": "informalité, accès au financement, droits d'auteur, débouchés limités"},
    {"code": "textile", "libelle": "Textile, habillement et mode",
     "sous_secteurs": ["Confection", "Pagne et tissage", "Égrenage et filature", "Distribution de prêt-à-porter"],
     "indicateurs": "production de coton, importations de friperie et de textiles, AGOA",
     "risques": "concurrence asiatique et friperie, coût de l'énergie, tendances de la mode"},
    {"code": "numerique_ecommerce", "libelle": "Services numériques et e-commerce",
     "sous_secteurs": ["Plateforme e-commerce", "Marketplace", "Services en ligne", "Logistique e-commerce"],
     "indicateurs": "taux de pénétration d'internet et du mobile money, achats en ligne, logistique du dernier kilomètre",
     "risques": "confiance des consommateurs, logistique, paiement à la livraison, concurrence des plateformes"},
    {"code": "dechets", "libelle": "Gestion des déchets et économie circulaire",
     "sous_secteurs": ["Collecte des déchets", "Recyclage des plastiques", "Valorisation énergétique",
                       "Compostage"],
     "indicateurs": "volumes de déchets produits, politiques publiques (ANAGED), interdiction des sachets plastiques",
     "risques": "dépendance aux contrats publics, débouchés des matières recyclées, réglementation"},
    {"code": "administration", "libelle": "Administration et services publics",
     "sous_secteurs": ["Prestations aux administrations", "Délégation de service public"],
     "indicateurs": "budget de l'État, programmes publics, commande publique",
     "risques": "délais de paiement, procédures de passation des marchés, changements de politique"},
    {"code": "autres", "libelle": "Autres secteurs", "sous_secteurs": [],
     "indicateurs": "taille et croissance du marché, acteurs, réglementation applicable",
     "risques": "marché, concurrence, financement, exploitation"},
]

PAR_CODE = {s["code"]: s for s in SECTEURS}


def trouver(code: str = "", libelle: str = "") -> Optional[Dict]:
    """Retrouve un secteur par code ou par libelle (dossiers anterieurs saisis en texte libre)."""
    if code and code in PAR_CODE:
        return PAR_CODE[code]
    if libelle:
        l = libelle.lower()
        for s in SECTEURS:
            if s["libelle"].lower() in l:
                return s
    return None


def reperes(code: str = "", libelle: str = "") -> str:
    """Axes d'analyse sectoriels transmis a l'IA (texte vide si secteur inconnu)."""
    s = trouver(code, libelle)
    if not s:
        return ""
    return (f"Secteur de référence : {s['libelle']}. Indicateurs à documenter en priorité : {s['indicateurs']}. "
            f"Risques typiques à examiner : {s['risques']}.")
