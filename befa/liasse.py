"""
Import des liasses fiscales SYSCOHADA (format DGI Cote d'Ivoire, "systeme
normal") : bilan actif, bilan passif, compte de resultat.

Chaine en trois temps :
  1. `selectionner_pages` : reperage des pages utiles dans le PDF (texte
     extrait par pypdf), pour n'envoyer a l'IA que ce qui sert ;
  2. `lire_liasse` : Claude lit ces pages et renvoie chaque ligne avec son
     code de reference (AD..BZ, CA..DZ, TA..XI) au format JSON impose ;
  3. `convertir` (deterministe, sans IA) : regroupement en bilan et compte de
     resultat condenses + CONTROLES contre les totaux imprimes dans la liasse
     (BZ, DZ, XB, XC, XI...). Tout ecart est remonte au consultant, qui valide
     ou corrige toujours les chiffres avant generation du Business Plan.
"""
import io
import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from .historique import BilanCondense, CompteResultatCondense, PosteImmobilisation
from .ia import IaError, appeler, bloc_pdf

# ----------------------------------------------------------------------
# Referentiel SYSCOHADA revise (liasse DGI, systeme normal)
# ----------------------------------------------------------------------
IMMOBILISATIONS = {
    "AE": "Frais de développement et de prospection",
    "AF": "Brevets, licences, logiciels et droits similaires",
    "AG": "Fonds commercial et droit au bail",
    "AH": "Autres immobilisations incorporelles",
    "AJ": "Terrains",
    "AK": "Bâtiments",
    "AL": "Aménagements, agencements et installations",
    "AM": "Matériel, mobilier et actifs biologiques",
    "AN": "Matériel de transport",
    "AP": "Avances et acomptes versés sur immobilisations",
    "AR": "Titres de participation",
    "AS": "Autres immobilisations financières",
}
TOTAUX_IMMOBILISATIONS = {"AD": "Immobilisations incorporelles", "AI": "Immobilisations corporelles",
                          "AQ": "Immobilisations financières"}
SOUS_LIGNES = {"AD": ["AE", "AF", "AG", "AH"], "AI": ["AJ", "AK", "AL", "AM", "AN", "AP"], "AQ": ["AR", "AS"]}

# Charges du compte de resultat (le signe imprime depend de l'editeur de la liasse)
CHARGES = {"RA", "RB", "RC", "RD", "RE", "RF", "RG", "RH", "RI", "RJ", "RK", "RL",
           "RM", "RN", "RO", "RP", "RQ", "RS"}
VARIATIONS_STOCKS = {"RB", "RD", "RF"}

TOLERANCE = 1.0  # ecart admis (FCFA) entre un total recalcule et le total imprime


# ----------------------------------------------------------------------
# 1. Selection des pages
# ----------------------------------------------------------------------
_MOTIFS_PAGES = {
    "actif": (re.compile(r"\bBILAN\b", re.I), re.compile(r"\bACTIF\b", re.I), re.compile(r"\bBZ\b")),
    "passif": (re.compile(r"\bBILAN\b", re.I), re.compile(r"\bPASSIF\b", re.I), re.compile(r"\bDZ\b")),
    "resultat": (re.compile(r"COMPTE DE R[ÉE]SULTAT", re.I), re.compile(r"\bXI\b")),
}


def selectionner_pages(pdf: bytes) -> Tuple[bytes, List[int]]:
    """Retourne (pdf_reduit, numeros_de_pages_1_based).

    Garde la page d'identification (1re page) et les pages du bilan et du
    compte de resultat. Si rien n'est reconnu (PDF scanne, sans texte), le
    document complet est renvoye : l'IA lit alors les images.
    """
    from pypdf import PdfReader, PdfWriter

    lecteur = PdfReader(io.BytesIO(pdf))
    retenues: List[int] = []
    for i, page in enumerate(lecteur.pages):
        try:
            texte = page.extract_text() or ""
        except Exception:
            texte = ""
        for motifs in _MOTIFS_PAGES.values():
            if all(m.search(texte) for m in motifs):
                retenues.append(i)
                break
    if not retenues:
        return pdf, list(range(1, len(lecteur.pages) + 1))
    pages = sorted(set([0] + retenues))
    ecrivain = PdfWriter()
    for i in pages:
        ecrivain.add_page(lecteur.pages[i])
    sortie = io.BytesIO()
    ecrivain.write(sortie)
    return sortie.getvalue(), [i + 1 for i in pages]


# ----------------------------------------------------------------------
# 2. Lecture par l'IA
# ----------------------------------------------------------------------
_NOMBRE = {"type": ["number", "null"]}
SCHEMA_LIASSE = {
    "type": "object",
    "properties": {
        "entreprise": {
            "type": "object",
            "properties": {
                "raison_sociale": {"type": "string"},
                "sigle": {"type": "string"},
                "ncc": {"type": "string"},
                "adresse": {"type": "string"},
                "exercice_clos_le": {"type": "string"},
                "duree_mois": {"type": ["integer", "null"]},
            },
            "required": ["raison_sociale", "sigle", "ncc", "adresse", "exercice_clos_le", "duree_mois"],
            "additionalProperties": False,
        },
        "actif": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {"ref": {"type": "string"}, "libelle": {"type": "string"},
                               "brut": _NOMBRE, "amortissements": _NOMBRE, "net": _NOMBRE, "net_n1": _NOMBRE},
                "required": ["ref", "libelle", "brut", "amortissements", "net", "net_n1"],
                "additionalProperties": False,
            },
        },
        "passif": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {"ref": {"type": "string"}, "libelle": {"type": "string"}, "n": _NOMBRE, "n1": _NOMBRE},
                "required": ["ref", "libelle", "n", "n1"],
                "additionalProperties": False,
            },
        },
        "compte_resultat": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {"ref": {"type": "string"}, "libelle": {"type": "string"}, "n": _NOMBRE, "n1": _NOMBRE},
                "required": ["ref", "libelle", "n", "n1"],
                "additionalProperties": False,
            },
        },
        "remarques": {"type": "string"},
    },
    "required": ["entreprise", "actif", "passif", "compte_resultat", "remarques"],
    "additionalProperties": False,
}

CONSIGNE_LECTURE = """Tu lis une liasse fiscale ivoirienne au format SYSCOHADA révisé (système normal, DGI).
Recopie fidèlement, sans rien calculer ni corriger :

1. "entreprise" : raison sociale (ou nom de l'exploitant), sigle, numéro de compte contribuable (NCC),
   adresse, date de clôture de l'exercice (JJ/MM/AAAA), durée de l'exercice en mois.
2. "actif" : chaque ligne du BILAN ACTIF portant un code REF (AD à BZ), y compris les totaux
   (AD, AI, AQ, AZ, BK, BT, BZ). Colonnes : brut, amortissements et dépréciations, net de l'exercice N,
   net de l'exercice N-1.
3. "passif" : chaque ligne du BILAN PASSIF portant un code REF (CA à DZ), y compris les totaux
   (CP, DD, DF, DP, DT, DZ). Colonnes : exercice N et exercice N-1.
4. "compte_resultat" : chaque ligne du COMPTE DE RÉSULTAT portant un code REF (TA à XI), y compris
   les soldes (XA, XB, XC, XD, XE, XF, XG, XH, XI). Colonnes : exercice N et exercice N-1.

Règles :
- Montants en francs CFA, nombres entiers sans séparateur de milliers (« 12 600 000 » -> 12600000).
- Garde le signe tel qu'il est imprimé : un montant entre parenthèses ou précédé de « - » est négatif.
- Cellule vide ou tiret seul -> null. N'invente aucune ligne absente du document.
- Ne confonds pas la colonne « NOTE » (petit numéro de note annexe) avec un montant.
- Le code REF est toujours de deux lettres majuscules.
- "remarques" : signale brièvement toute difficulté de lecture (page illisible, colonne ambiguë), sinon "".
"""


@dataclass
class LiasseLue:
    entreprise: Dict[str, object]
    actif: Dict[str, Dict[str, Optional[float]]]  # ref -> {brut, amortissements, net, net_n1, libelle}
    passif: Dict[str, Dict[str, Optional[float]]]  # ref -> {n, n1, libelle}
    compte_resultat: Dict[str, Dict[str, Optional[float]]]  # ref -> {n, n1, libelle}
    remarques: str = ""
    pages: List[int] = field(default_factory=list)

    @staticmethod
    def depuis_json(d: dict, pages: Optional[List[int]] = None) -> "LiasseLue":
        def indexer(lignes):
            res = {}
            for l in lignes:
                ref = str(l.get("ref", "")).strip().upper()
                if re.fullmatch(r"[A-Z]{2}", ref):
                    res[ref] = {k: v for k, v in l.items() if k != "ref"}
            return res

        return LiasseLue(
            entreprise=d.get("entreprise", {}),
            actif=indexer(d.get("actif", [])),
            passif=indexer(d.get("passif", [])),
            compte_resultat=indexer(d.get("compte_resultat", [])),
            remarques=d.get("remarques", "") or "",
            pages=pages or [],
        )

    def vers_json(self) -> dict:
        return {
            "entreprise": self.entreprise,
            "actif": [{"ref": k, **v} for k, v in self.actif.items()],
            "passif": [{"ref": k, **v} for k, v in self.passif.items()],
            "compte_resultat": [{"ref": k, **v} for k, v in self.compte_resultat.items()],
            "remarques": self.remarques,
            "pages": self.pages,
        }


def lire_liasse(pdf: bytes, nom_fichier: str = "liasse.pdf") -> LiasseLue:
    """Lit une liasse SYSCOHADA (PDF) avec Claude. Leve IaError en cas d'echec."""
    reduit, pages = selectionner_pages(pdf)
    data = appeler(
        [bloc_pdf(reduit, nom_fichier), {"type": "text", "text": CONSIGNE_LECTURE}],
        schema=SCHEMA_LIASSE,
        effort="medium",
    )
    liasse = LiasseLue.depuis_json(data, pages)
    if not (liasse.actif or liasse.passif or liasse.compte_resultat):
        raise IaError("Aucune ligne de bilan ni de compte de résultat n'a été reconnue dans ce document : "
                      "vérifiez qu'il s'agit bien d'une liasse SYSCOHADA.")
    return liasse


# ----------------------------------------------------------------------
# 3. Conversion deterministe + controles
# ----------------------------------------------------------------------
def _v(table: Dict[str, Dict[str, Optional[float]]], ref: str, col: str) -> float:
    ligne = table.get(ref)
    if not ligne:
        return 0.0
    val = ligne.get(col)
    return float(val) if val is not None else 0.0


def _present(table, ref: str, col: str) -> bool:
    ligne = table.get(ref)
    return bool(ligne) and ligne.get(col) is not None


def _somme(table, refs, col, total: Optional[str] = None) -> float:
    """Somme des sous-lignes ; a defaut, le total imprime."""
    if any(_present(table, r, col) for r in refs):
        return sum(_v(table, r, col) for r in refs)
    return _v(table, total, col) if total else 0.0


def _fmt(x: float) -> str:
    return f"{x:,.0f}".replace(",", " ")


def _annee(date_cloture: str) -> Optional[int]:
    m = re.search(r"(19|20)\d\d", date_cloture or "")
    return int(m.group(0)) if m else None


def _controle(alertes: List[str], libelle: str, calcule: float, table, ref: str, col: str):
    if _present(table, ref, col):
        imprime = _v(table, ref, col)
        if abs(calcule - imprime) > TOLERANCE:
            alertes.append(f"{libelle} : total recalculé {_fmt(calcule)} ≠ total imprimé {ref} {_fmt(imprime)} "
                           f"(écart {_fmt(calcule - imprime)}). Vérifiez les lignes concernées.")


def _bilan(l: LiasseLue, colonne_n: bool, exercice: str, alertes: List[str]) -> BilanCondense:
    A, P = l.actif, l.passif
    col_net = "net" if colonne_n else "net_n1"
    col_p = "n" if colonne_n else "n1"

    immos: List[PosteImmobilisation] = []
    for total, refs in SOUS_LIGNES.items():
        presentes = [r for r in refs if _present(A, r, col_net) or (colonne_n and _present(A, r, "brut"))]
        if not presentes and (_present(A, total, col_net) or (colonne_n and _present(A, total, "brut"))):
            presentes = [total]
        for r in presentes:
            lib = IMMOBILISATIONS.get(r) or TOTAUX_IMMOBILISATIONS[r]
            net = _v(A, r, col_net)
            if colonne_n and _present(A, r, "brut"):
                brut = _v(A, r, "brut")
                amort = abs(_v(A, r, "amortissements"))
                if abs((brut - amort) - net) > TOLERANCE and _present(A, r, "net"):
                    alertes.append(f"Ligne {r} ({lib}) : brut - amortissements ({_fmt(brut - amort)}) ≠ net ({_fmt(net)}).")
            else:
                brut, amort = net, 0.0
            if brut or amort or net:
                immos.append(PosteImmobilisation(lib, brut, amort))

    b = BilanCondense(
        exercice=exercice,
        immobilisations=immos,
        detail_amortissements=colonne_n,
        stocks=_v(A, "BB", col_net),
        avances_fournisseurs=_v(A, "BH", col_net),
        creances_clients=_v(A, "BI", col_net),
        autres_creances=_v(A, "BJ", col_net) + _v(A, "BA", col_net) + _v(A, "BU", col_net),
        tresorerie_actif=_somme(A, ["BQ", "BR", "BS"], col_net, "BT"),
        capital=_v(P, "CA", col_p) - abs(_v(P, "CB", col_p)),
        reserves_report=sum(_v(P, r, col_p) for r in ("CD", "CE", "CF", "CG", "CH", "CM")),
        resultat_net=_v(P, "CJ", col_p),
        subventions_investissement=_v(P, "CL", col_p),
        dettes_financieres=_somme(P, ["DA", "DB", "DC"], col_p, "DD"),
        fournisseurs=_v(P, "DJ", col_p),
        dettes_fiscales_sociales=_v(P, "DK", col_p),
        autres_dettes=sum(_v(P, r, col_p) for r in ("DH", "DI", "DM", "DN", "DV")),
        tresorerie_passif=_somme(P, ["DQ", "DR"], col_p, "DT"),
    )
    # stocks / creances : si les sous-lignes manquent mais le total BK existe
    if not any(_present(A, r, col_net) for r in ("BA", "BB", "BH", "BI", "BJ")) and _present(A, "BK", col_net):
        b.autres_creances = _v(A, "BK", col_net)

    _controle(alertes, f"Actif immobilisé {exercice}", b.actif_immobilise_net, A, "AZ", col_net)
    _controle(alertes, f"Actif circulant {exercice}", b.actif_circulant - _v(A, "BU", col_net), A, "BK", col_net)
    _controle(alertes, f"Total actif {exercice}", b.total_actif, A, "BZ", col_net)
    _controle(alertes, f"Capitaux propres {exercice}", b.capitaux_propres, P, "CP", col_p)
    _controle(alertes, f"Total passif {exercice}", b.total_passif, P, "DZ", col_p)
    return b


def _compte_resultat(l: LiasseLue, col: str, exercice: str, alertes: List[str]) -> CompteResultatCondense:
    T = l.compte_resultat
    v = lambda r: _v(T, r, col)
    produits_va = sum(v(r) for r in ("TE", "TF", "TG", "TH", "TI"))
    ca = v("XB") if _present(T, "XB", col) else sum(v(r) for r in ("TA", "TB", "TC", "TD"))
    if _present(T, "XB", col) and any(_present(T, r, col) for r in ("TA", "TB", "TC", "TD")):
        _controle(alertes, f"Chiffre d'affaires {exercice}", sum(v(r) for r in ("TA", "TB", "TC", "TD")), T, "XB", col)

    # Convention de signe des charges : on retient celle qui redonne la valeur
    # ajoutee imprimee (XC).
    def charge_factory(signe_charges: int, signe_variations: int):
        def c(r: str) -> float:
            return v(r) * (signe_variations if r in VARIATIONS_STOCKS else signe_charges)
        return c

    def va(c) -> float:
        return (ca + produits_va - sum(c(r) for r in ("RA", "RB", "RC", "RD", "RE", "RF", "RG", "RH", "RI", "RJ")))

    hypotheses = [(-1, -1), (1, 1), (1, -1)]
    c = None
    if _present(T, "XC", col):
        for sc, sv in hypotheses:
            cand = charge_factory(sc, sv)
            if abs(va(cand) - v("XC")) <= TOLERANCE:
                c = cand
                break
        if c is None:
            alertes.append(f"Compte de résultat {exercice} : la valeur ajoutée recalculée ne correspond pas au "
                           f"total imprimé XC ({_fmt(v('XC'))}). Vérifiez les lignes de charges (RA à RJ).")
    if c is None:
        negatives = sum(1 for r in CHARGES - VARIATIONS_STOCKS if v(r) < 0)
        positives = sum(1 for r in CHARGES - VARIATIONS_STOCKS if v(r) > 0)
        c = charge_factory(-1, -1) if negatives > positives else charge_factory(1, 1)

    cr = CompteResultatCondense(
        exercice=exercice,
        chiffre_affaires=ca,
        achats_matieres=c("RA") + c("RB") + c("RC") + c("RD"),
        autres_achats=c("RE") + c("RF"),
        transport=c("RG"),
        services_exterieurs=c("RH"),
        impots_taxes=c("RI"),
        autres_charges=c("RJ") - produits_va,
        salaires=c("RK"),
        dotations=c("RL") - v("TJ"),
        frais_financiers=c("RM") + c("RN") - (v("TK") + v("TL") + v("TM")),
        resultat_hao=v("TN") + v("TO") - c("RO") - c("RP"),
        impot_sur_benefice=c("RQ") + c("RS"),
    )
    _controle(alertes, f"Excédent brut d'exploitation {exercice}", cr.ebe, T, "XD", col)
    _controle(alertes, f"Résultat d'exploitation {exercice}", cr.resultat_exploitation, T, "XE", col)
    _controle(alertes, f"Résultat net {exercice}", cr.resultat_net, T, "XI", col)
    return cr


@dataclass
class ImportLiasse:
    entreprise: Dict[str, object]
    bilan_n: BilanCondense
    bilan_n1: Optional[BilanCondense]
    cr_n: CompteResultatCondense
    cr_n1: Optional[CompteResultatCondense]
    alertes: List[str]
    remarques: str
    pages: List[int]


def _colonne_vide(table, cols) -> bool:
    return not any(ligne.get(col) not in (None, 0) for ligne in table.values() for col in cols)


def convertir(l: LiasseLue) -> ImportLiasse:
    """Regroupe une liasse lue en bilan et compte de resultat condenses, avec controles."""
    alertes: List[str] = []
    annee = _annee(str(l.entreprise.get("exercice_clos_le", "")))
    ex_n = str(annee) if annee else "N"
    ex_n1 = str(annee - 1) if annee else "N-1"

    bilan_n = _bilan(l, True, ex_n, alertes)
    cr_n = _compte_resultat(l, "n", ex_n, alertes)

    # Exercice N-1 : absent d'un bilan d'ouverture (entreprise nouvelle)
    avec_n1 = not (_colonne_vide(l.actif, ["net_n1"]) and _colonne_vide(l.passif, ["n1"])
                   and _colonne_vide(l.compte_resultat, ["n1"]))
    bilan_n1 = _bilan(l, False, ex_n1, alertes) if avec_n1 else None
    cr_n1 = _compte_resultat(l, "n1", ex_n1, alertes) if avec_n1 else None

    for b, cr in ((bilan_n, cr_n), (bilan_n1, cr_n1)):
        if b and cr and abs(b.resultat_net - cr.resultat_net) > TOLERANCE:
            alertes.append(f"Exercice {b.exercice} : résultat net du bilan (CJ = {_fmt(b.resultat_net)}) différent de "
                           f"celui du compte de résultat ({_fmt(cr.resultat_net)}).")
    # Pratique BEFA : les avances des associes (compte courant) sont classees en
    # capitaux propres. La liasse les range souvent en "Autres dettes" (DM) :
    # on le signale, sans reclasser d'office.
    for b in (bilan_n, bilan_n1):
        if b and b.total_passif and _v(l.passif, "DM", "n" if b is bilan_n else "n1") > 0.2 * b.total_passif:
            alertes.append(
                f"Exercice {b.exercice} : « Autres dettes » (DM) représentent "
                f"{_v(l.passif, 'DM', 'n' if b is bilan_n else 'n1') / b.total_passif:.0%} du bilan. S'il s'agit "
                "d'avances des associés en compte courant, reclassez-les en « Compte courant d'associés » "
                "(capitaux propres), comme dans les dossiers BEFA."
            )
    duree = l.entreprise.get("duree_mois")
    if duree and duree != 12:
        alertes.append(f"Exercice de {duree} mois : les comparaisons N / N-1 sont à interpréter avec prudence.")

    return ImportLiasse(l.entreprise, bilan_n, bilan_n1, cr_n, cr_n1, alertes, l.remarques, l.pages)
