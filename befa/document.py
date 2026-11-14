"""
Business Plan au format BEFA Consulting : document Word (python-docx) puis
PDF (LibreOffice). Lot 5.

Structure calquee sur les modeles du cabinet :
  - nouveau projet       : modele COREAL FOOD (7 parties + conclusion) ;
  - entreprise existante : modele O'IVOIRE (4 parties + conclusion).

Table des matieres en deux passes : LibreOffice ne met pas a jour les
numeros de page d'un sommaire Word. On genere un premier PDF, on y repere
la page de chaque titre, puis on regenere le document avec les vrais
numeros (et une 3e passe si la pagination a bouge).
"""
import io
import os
import re
import tempfile
import unicodedata
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence, Tuple

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from . import commentaires as C
from . import graphiques, trends
from .dossier import DossierBEFA
from .pdf import docx_vers_pdf
from .projection import Ligne, Projection, calculer_projection
from .redaction import Redaction, commentaires_financiers

RACINE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOGO = os.path.join(RACINE, "assets", "branding", "logo_befa.png")

ORANGE = RGBColor(0xF2, 0x7F, 0x24)
ORANGE_FONCE = RGBColor(0xB8, 0x5A, 0x12)
GRIS = RGBColor(0x5B, 0x61, 0x68)
ENCRE = RGBColor(0x1F, 0x23, 0x28)
HEX_ORANGE, HEX_ORANGE_CLAIR, HEX_GRIS_CLAIR, HEX_BORDURE = "F27F24", "FCE6D4", "F1F0F0", "C9C6C7"
POLICE = "Arial"

CABINET = "Bureau d'Études Financières et d'Assistance-Consulting"
CONTACTS = ["Tél. : 27 22 35 60 94 / 07 58 77 42 67", "infos@befaconsulting.org — www.befaconsulting.com",
            "23 BP 967 Abidjan 23 — Cocody Riviera Faya"]
MARQUE_SOMMAIRE = "§§"  # numero de page provisoire (1re passe) : repere les pages du sommaire


# ----------------------------------------------------------------------
# Mise en forme de base
# ----------------------------------------------------------------------
def _nombre(v, style: str = "") -> str:
    if v is None:
        return ""
    if style == "ratio":
        return f"{v:.1f}".replace(".", ",").replace("-", "−")
    if style == "pct":
        return f"{v * 100:.0f} %".replace("-", "−")
    if abs(v) < 0.5:
        return "–"
    return f"{v:,.0f}".replace(",", " ").replace("-", "−")


def _ombre(cellule, couleur: str):
    tc = cellule._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), couleur)
    tc.append(shd)


def _bordures(table, couleur: str = HEX_BORDURE):
    tbl_pr = table._tbl.tblPr
    bords = OxmlElement("w:tblBorders")
    for cote in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = OxmlElement(f"w:{cote}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "4")
        b.set(qn("w:color"), couleur)
        bords.append(b)
    tbl_pr.append(bords)


def _grille(table, largeurs_cm):
    """Largeurs reelles des colonnes : LibreOffice lit la grille du tableau (w:tblGrid),
    pas la largeur posee sur chaque cellule."""
    table.autofit = False
    for col, l in zip(table.columns, largeurs_cm):
        col.width = Cm(l)
    tbl_pr = table._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tbl_pr.append(layout)
    marges = OxmlElement("w:tblCellMar")  # marges interieures reduites : 0,08 cm a gauche et a droite
    for cote, val in (("left", 45), ("right", 45), ("top", 15), ("bottom", 15)):
        m = OxmlElement(f"w:{cote}")
        m.set(qn("w:w"), str(val))
        m.set(qn("w:type"), "dxa")
        marges.append(m)
    tbl_pr.append(marges)


def _champ(run, code: str):
    """Insere un champ Word (ex. PAGE) dans un run."""
    for typ, texte in (("begin", None), (None, code), ("separate", None), (None, "1"), ("end", None)):
        if typ:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), typ)
        elif texte == code:
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = f" {code} "
        else:
            el = OxmlElement("w:t")
            el.text = texte
        run._r.append(el)


def _bordure_paragraphe(p, cote: str = "bottom", couleur: str = HEX_ORANGE, epaisseur: int = 8):
    ppr = p._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    b = OxmlElement(f"w:{cote}")
    b.set(qn("w:val"), "single")
    b.set(qn("w:sz"), str(epaisseur))
    b.set(qn("w:space"), "4")
    b.set(qn("w:color"), couleur)
    pbdr.append(b)
    ppr.append(pbdr)


def _normaliser(t: str) -> str:
    t = unicodedata.normalize("NFKC", t).lower()
    return re.sub(r"[\W_]+", "", t)


@dataclass
class EntreeSommaire:
    niveau: int
    titre: str


class Redacteur:
    """Construit le document Word du Business Plan."""

    def __init__(self, pages: Optional[Dict[int, int]] = None):
        self.doc = Document()
        self.pages = pages  # index de l'entree du sommaire -> numero de page
        self.sommaire: List[EntreeSommaire] = []
        self.figures = 0
        self.renvois: Dict[str, int] = {}
        self._styles()

    # --- styles ---------------------------------------------------------
    def _styles(self):
        st = self.doc.styles
        normal = st["Normal"]
        normal.font.name = POLICE
        normal.font.size = Pt(10.5)
        normal.font.color.rgb = ENCRE
        normal.element.rPr.rFonts.set(qn("w:eastAsia"), POLICE)
        pf = normal.paragraph_format
        pf.space_after = Pt(6)
        pf.line_spacing = 1.15
        for nom, taille, couleur, avant, apres in (("Heading 1", 16, ORANGE, 0, 14), ("Heading 2", 13, ENCRE, 14, 6),
                                                   ("Heading 3", 11, ORANGE_FONCE, 10, 4)):
            s = st[nom]
            s.font.name = POLICE
            polices = s.element.rPr.rFonts
            polices.set(qn("w:eastAsia"), POLICE)
            for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
                if polices.get(qn(attr)) is not None:  # la police du theme Word l'emporterait sur Arial
                    del polices.attrib[qn(attr)]
            s.font.size = Pt(taille)
            s.font.bold = True
            s.font.italic = False
            s.font.color.rgb = couleur
            s.paragraph_format.space_before = Pt(avant)
            s.paragraph_format.space_after = Pt(apres)
            s.paragraph_format.keep_with_next = True
        sec = self.doc.sections[0]
        sec.page_height, sec.page_width = Cm(29.7), Cm(21)
        sec.left_margin = sec.right_margin = Cm(1.9)
        sec.top_margin, sec.bottom_margin = Cm(2.0), Cm(1.8)
        sec.header_distance, sec.footer_distance = Cm(0.8), Cm(0.8)
        sec.different_first_page_header_footer = True

    @property
    def largeur(self) -> float:
        s = self.doc.sections[0]
        return (s.page_width - s.left_margin - s.right_margin) / 360000  # cm

    # --- en-tete / pied de page ---------------------------------------------
    def entete_pied(self):
        sec = self.doc.sections[0]
        h = sec.header.paragraphs[0]
        h.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        if os.path.exists(LOGO):
            h.add_run().add_picture(LOGO, height=Cm(0.9))
        f = sec.footer.paragraphs[0]
        _bordure_paragraphe(f, "top", HEX_ORANGE, 6)
        f.paragraph_format.tab_stops.add_tab_stop(Cm(self.largeur), WD_TAB_ALIGNMENT.RIGHT)
        r = f.add_run(f"Réalisé par {CABINET}")
        r.font.size, r.font.color.rgb = Pt(8), GRIS
        r = f.add_run("\tPage ")
        r.font.size, r.font.color.rgb = Pt(8), GRIS
        r = f.add_run()
        r.font.size, r.font.color.rgb = Pt(8), GRIS
        _champ(r, "PAGE")

    # --- blocs ---------------------------------------------------------------
    def titre(self, texte: str, niveau: int, saut: bool = False):
        if saut or niveau == 1:
            self.saut_page()
        self.doc.add_heading(texte, level=niveau)
        if niveau <= 2:
            self.sommaire.append(EntreeSommaire(niveau, texte))

    def saut_page(self):
        paras = self.doc.paragraphs
        if paras and not paras[-1].text.strip() and not paras[-1].runs:
            paras[-1].add_run().add_break(WD_BREAK.PAGE)
        else:
            self.doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    def paragraphe(self, texte: str, italique=False, taille=None, couleur=None, aligne=WD_ALIGN_PARAGRAPH.JUSTIFY):
        p = self.doc.add_paragraph()
        p.alignment = aligne
        self._texte(p, texte, italique, taille, couleur)
        return p

    def _texte(self, p, texte, italique=False, taille=None, couleur=None, gras=False):
        """Ecrit un texte ; les renvois [Sx] deviennent des appels de note [n] en exposant."""
        for morceau in re.split(r"(\[S\d+\])", texte):
            if not morceau:
                continue
            m = re.fullmatch(r"\[(S\d+)\]", morceau)
            if m:
                n = self.renvois.get(m.group(1))
                if n is None:
                    continue
                r = p.add_run(f"[{n}]")
                r.font.superscript = True
                r.font.color.rgb = ORANGE_FONCE
                continue
            r = p.add_run(morceau)
            r.italic, r.bold = italique, gras
            if taille:
                r.font.size = Pt(taille)
            if couleur:
                r.font.color.rgb = couleur

    def puces(self, elements: Sequence[str]):
        for e in elements:
            p = self.doc.add_paragraph(style="List Bullet")
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.space_after = Pt(3)
            # "Libelle : texte" -> libelle en gras
            m = re.match(r"^([^:]{3,60}) : (.+)$", e)
            if m:
                self._texte(p, m.group(1) + " : ", gras=True)
                self._texte(p, m.group(2))
            else:
                self._texte(p, e)

    def a_completer(self, quoi: str):
        self.paragraphe(f"[À compléter par le consultant : {quoi}]", italique=True, couleur=ORANGE_FONCE)

    def note(self, texte: str):
        p = self.paragraphe(texte, italique=True, taille=8.5, couleur=GRIS)
        p.paragraph_format.space_after = Pt(4)

    def figure(self, png: bytes, legende: str, largeur_cm: float = 16):
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.keep_with_next = True
        p.add_run().add_picture(io.BytesIO(png), width=Cm(min(largeur_cm, self.largeur)))
        self.figures += 1
        cap = self.paragraphe(f"Figure {self.figures} : {legende}", italique=True, taille=8.5, couleur=GRIS,
                              aligne=WD_ALIGN_PARAGRAPH.CENTER)
        cap.paragraph_format.space_after = Pt(10)

    def tableau(self, entetes: Sequence[str], lignes: Sequence[Ligne], largeur_libelle: float = 5.0,
                styles_colonnes: Optional[Sequence[str]] = None, taille: float = 8, titre: Optional[str] = None):
        """Tableau financier : 1re colonne libelle, colonnes suivantes numeriques."""
        if titre:
            p = self.paragraphe(titre, taille=9, couleur=ORANGE_FONCE)
            p.runs[0].bold = True
            p.paragraph_format.keep_with_next = True
            p.paragraph_format.space_after = Pt(3)
        ncol = len(entetes)
        if ncol >= 8:  # plan de financement avec annee 0 : tableau tres large
            taille, largeur_libelle = min(taille, 6.8), min(largeur_libelle, 3.3)
        t = self.doc.add_table(rows=1, cols=ncol)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        _bordures(t)
        larg_val = (self.largeur - largeur_libelle) / max(1, ncol - 1)
        largeurs = [largeur_libelle] + [larg_val] * (ncol - 1)
        _grille(t, largeurs)
        for i, e in enumerate(entetes):
            c = t.rows[0].cells[i]
            c.width = Cm(largeurs[i])
            _ombre(c, HEX_ORANGE)
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if i == 0 else WD_ALIGN_PARAGRAPH.RIGHT
            r = p.add_run(e)
            r.bold, r.font.size, r.font.color.rgb = True, Pt(taille), RGBColor(0xFF, 0xFF, 0xFF)
        self._repeter_entete(t.rows[0])
        for l in lignes:
            cells = t.add_row().cells
            for i in range(ncol):
                c = cells[i]
                c.width = Cm(largeurs[i])
                p = c.paragraphs[0]
                p.paragraph_format.space_after = Pt(0)
                if i == 0:
                    texte = l.libelle
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    if l.style == "detail":
                        p.paragraph_format.left_indent = Cm(0.25)
                else:
                    v = l.valeurs[i - 1] if i - 1 < len(l.valeurs) else None
                    style_col = styles_colonnes[i - 1] if styles_colonnes else ""
                    texte = _nombre(v, "ratio" if l.style == "ratio" else style_col)
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                r = p.add_run(texte)
                r.font.size = Pt(taille)
                r.bold = l.style in ("total", "sous_total")
                r.italic = l.style == "detail"
                if l.style == "total":
                    _ombre(c, HEX_ORANGE_CLAIR)
                elif l.style == "sous_total":
                    _ombre(c, HEX_GRIS_CLAIR)
        self.doc.add_paragraph().paragraph_format.space_after = Pt(2)
        return t

    def tableau_texte(self, entetes: Optional[Sequence[str]], lignes: Sequence[Sequence[str]],
                      largeurs: Sequence[float], taille: float = 9, gras_premiere_colonne: bool = False):
        t = self.doc.add_table(rows=0, cols=len(largeurs))
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        _bordures(t)
        _grille(t, largeurs)
        if entetes:
            cells = t.add_row().cells
            for i, e in enumerate(entetes):
                cells[i].width = Cm(largeurs[i])
                _ombre(cells[i], HEX_ORANGE)
                r = cells[i].paragraphs[0].add_run(e)
                r.bold, r.font.size, r.font.color.rgb = True, Pt(taille), RGBColor(0xFF, 0xFF, 0xFF)
            self._repeter_entete(t.rows[0])
        for ligne in lignes:
            cells = t.add_row().cells
            for i, v in enumerate(ligne):
                cells[i].width = Cm(largeurs[i])
                p = cells[i].paragraphs[0]
                p.paragraph_format.space_after = Pt(0)
                r = p.add_run(str(v))
                r.font.size = Pt(taille)
                if gras_premiere_colonne and i == 0:
                    r.bold = True
                    _ombre(cells[i], HEX_GRIS_CLAIR)
        self.doc.add_paragraph().paragraph_format.space_after = Pt(2)
        return t

    @staticmethod
    def _repeter_entete(ligne):
        trpr = ligne._tr.get_or_add_trPr()
        el = OxmlElement("w:tblHeader")
        el.set(qn("w:val"), "true")
        trpr.append(el)

    # --- page de garde et sommaire -------------------------------------------
    def page_de_garde(self, d: DossierBEFA, titre_doc: str):
        e, p = d.entreprise, d.projet
        for _ in range(5):
            self.doc.add_paragraph()
        for texte, taille, couleur, gras in ((e.raison_sociale.upper(), 24, ENCRE, True),
                                             (titre_doc, 30, ORANGE, True),
                                             (p.intitule, 15, GRIS, False),
                                             (d.annee_document, 16, ENCRE, True)):
            if not texte:
                continue
            q = self.doc.add_paragraph()
            q.alignment = WD_ALIGN_PARAGRAPH.CENTER
            q.paragraph_format.space_after = Pt(10)
            r = q.add_run(texte)
            r.font.size, r.font.color.rgb, r.bold = Pt(taille), couleur, gras
        sep = self.doc.add_paragraph()
        _bordure_paragraphe(sep, "bottom", HEX_ORANGE, 12)
        for _ in range(6):
            self.doc.add_paragraph()
        q = self.doc.add_paragraph()
        q.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = q.add_run("Réalisé par :")
        r.font.size, r.font.color.rgb, r.italic = Pt(11), GRIS, True
        if os.path.exists(LOGO):
            q = self.doc.add_paragraph()
            q.alignment = WD_ALIGN_PARAGRAPH.CENTER
            q.add_run().add_picture(LOGO, width=Cm(8.5))
        for c in CONTACTS:
            q = self.doc.add_paragraph()
            q.alignment = WD_ALIGN_PARAGRAPH.CENTER
            q.paragraph_format.space_after = Pt(0)
            r = q.add_run(c)
            r.font.size, r.font.color.rgb = Pt(9), GRIS

    def emplacement_sommaire(self):
        self.saut_page()
        q = self.doc.add_paragraph()
        r = q.add_run("Table des matières")
        r.bold, r.font.size, r.font.color.rgb = True, Pt(16), ORANGE
        q.paragraph_format.space_after = Pt(12)
        self._ancre_sommaire = self.doc.add_paragraph()

    def remplir_sommaire(self):
        ancre = self._ancre_sommaire
        for i, e in enumerate(self.sommaire):
            p = ancre.insert_paragraph_before()
            pf = p.paragraph_format
            pf.tab_stops.add_tab_stop(Cm(self.largeur), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
            pf.space_after = Pt(2 if e.niveau == 2 else 4)
            pf.space_before = Pt(0 if e.niveau == 2 else 6)
            if e.niveau == 2:
                pf.left_indent = Cm(0.7)
            numero = str(self.pages.get(i, "")) if self.pages is not None else MARQUE_SOMMAIRE
            r = p.add_run(f"{e.titre}\t{numero}")
            r.font.size = Pt(10 if e.niveau == 1 else 9.5)
            r.bold = e.niveau == 1
            if e.niveau == 1:
                r.font.color.rgb = ORANGE_FONCE


# ----------------------------------------------------------------------
# Contenu
# ----------------------------------------------------------------------
def _section_ia(R: Redacteur, red: Optional[Redaction], cle: str, titre: Optional[str], niveau: int = 2,
                quoi: str = ""):
    s = red.section(cle) if red else None
    if titre:
        R.titre(titre, niveau)
    if not s or not (s.paragraphes or s.puces):
        R.a_completer(quoi or (s.titre if s else titre or cle))
        return
    for p in s.paragraphes:
        R.paragraphe(p)
    if s.puces:
        R.puces(s.puces)


def _commentaires(R: Redacteur, coms: Dict[str, List[str]], cle: str):
    for p in coms.get(cle, []):
        R.paragraphe(p)


def _swot(R: Redacteur, red: Optional[Redaction]):
    if not red or not any(red.swot.values()):
        R.a_completer("analyse SWOT")
        return
    sw = red.swot
    largeur = R.largeur / 2
    t = R.doc.add_table(rows=2, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    _bordures(t)
    _grille(t, [largeur, largeur])
    for (ligne, col), (cle, titre, fond) in zip(
            ((0, 0), (0, 1), (1, 0), (1, 1)),
            (("forces", "Forces", "E3F1E6"), ("faiblesses", "Faiblesses", "FBE3E1"),
             ("opportunites", "Opportunités", "E2ECF7"), ("menaces", "Menaces", "FCEBD9"))):
        c = t.rows[ligne].cells[col]
        c.width = Cm(largeur)
        _ombre(c, fond)
        p = c.paragraphs[0]
        r = p.add_run(titre)
        r.bold, r.font.size = True, Pt(10)
        for x in sw.get(cle, []):
            q = c.add_paragraph()
            q.paragraph_format.space_after = Pt(2)
            q.paragraph_format.left_indent = Cm(0.3)
            R._texte(q, "• " + x, taille=8.5)
    R.doc.add_paragraph()


def _identite(R: Redacteur, d: DossierBEFA):
    e = d.entreprise
    R.titre("Identification de l'entreprise", 3)
    lignes = [(k, v) for k, v in (
        ("Raison sociale", e.raison_sociale), ("Sigle", e.sigle), ("Forme juridique", e.forme_juridique),
        ("Capital social", f"{C.entier(e.capital)} FCFA" if e.capital else ""), ("Date de création", e.date_creation),
        ("RCCM", e.rccm), ("N° de compte contribuable", e.ncc), ("Régime fiscal", e.regime_fiscal),
        ("Centre des impôts", e.centre_impots), ("Siège social", e.siege or e.ville), ("Adresse", e.adresse),
        ("Téléphone", e.telephone), ("E-mail", e.email)) if v]
    R.tableau_texte(None, lignes, [5.5, R.largeur - 5.5], gras_premiere_colonne=True)
    if e.activite:
        R.titre("Nature de l'activité", 3)
        R.paragraphe(e.activite)
    if e.objet_social:
        R.titre("Objet social", 3)
        R.paragraphe(f"{e.raison_sociale} a pour objet, en République de Côte d'Ivoire et dans les États parties au "
                     f"Traité OHADA :")
        R.puces(e.objet_social)


def _actionnariat(R: Redacteur, d: DossierBEFA):
    if not d.actionnaires:
        R.a_completer("répartition du capital")
        return
    capital = d.entreprise.capital or sum(a.montant for a in d.actionnaires)
    lignes = [(a.nom, a.nationalite, a.nature, C.entier(a.montant),
               C.pct(a.part(capital)) if a.part(capital) is not None else "") for a in d.actionnaires]
    lignes.append(("TOTAL", "", "", C.entier(sum(a.montant for a in d.actionnaires)), "100 %"))
    R.tableau_texte(["Associé / actionnaire", "Nationalité", "Nature", "Capital (FCFA)", "Part"], lignes,
                    [5.6, 2.6, 3.6, 3.2, R.largeur - 15.0])


def _management(R: Redacteur, d: DossierBEFA):
    if d.gouvernance:
        R.paragraphe(d.gouvernance)
    if d.dirigeants:
        R.tableau_texte(["Nom", "Fonction", "Profil"], [(x.nom, x.fonction, x.profil) for x in d.dirigeants],
                        [5.0, 4.0, R.largeur - 9.0])
    if not d.gouvernance and not d.dirigeants:
        R.a_completer("organes de direction et équipe dirigeante")


def _partenaires(R: Redacteur, liste, quoi: str):
    if not liste:
        R.a_completer(quoi)
        return
    R.tableau_texte(["Nom", "Mode de règlement", "Délai", "Localisation"],
                    [(x.nom, x.mode_paiement, x.delai, x.localisation) for x in liste],
                    [6.0, 3.8, 3.6, R.largeur - 13.4])


def _produits(R: Redacteur, d: DossierBEFA):
    if not d.produits:
        R.a_completer("produits et services")
        return
    for x in d.produits:
        q = R.doc.add_paragraph(style="List Bullet")
        R._texte(q, x.nom + (" : " if x.description else ""), gras=True)
        if x.description:
            R._texte(q, x.description)


def _prets(R: Redacteur, proj: Projection, coms):
    _commentaires(R, coms, "financement")
    for pr in proj.entree.prets:
        differe = " et ".join(x for x in (
            f"{pr.differe_total_mois} mois de différé total" if pr.differe_total_mois else "",
            f"{pr.differe_partiel_mois} mois de différé partiel" if pr.differe_partiel_mois else "") if x) or "Aucun"
        taux = f"{pr.taux_annuel * 100:.2f}".rstrip("0").rstrip(".").replace(".", ",") + " %"
        lignes = [("Objet", pr.objet or "—"), ("Montant", f"{C.entier(pr.montant)} FCFA"), ("Taux d'intérêt", taux),
                  ("Durée", f"{pr.duree_mois} mois"), ("Différé", differe),
                  ("Modalités", "Remboursement in fine" if pr.in_fine else
                   f"Remboursements {C.PERIODICITE.get(pr.periodicite, pr.periodicite)}")]
        if pr.garantie:
            lignes.append(("Garantie", pr.garantie))
        R.titre(f"{pr.libelle_type} — {pr.libelle}", 3)
        R.tableau_texte(None, lignes, [4.5, R.largeur - 4.5], gras_premiere_colonne=True)


def _investissements(R: Redacteur, proj: Projection):
    invs = proj.entree.investissements
    if not invs:
        R.a_completer("détail des investissements")
        return
    lignes = []
    for i in invs:
        lignes.append(Ligne(i.libelle, [i.montant]))
        for det in i.details:
            lignes.append(Ligne(det, [None], "detail"))
    lignes.append(Ligne("COÛT TOTAL DES INVESTISSEMENTS", [proj.cout_total_investissements], "total"))
    R.tableau(["Rubrique", "Montant (FCFA HT)"], lignes, largeur_libelle=R.largeur - 4.5, taille=9)
    total = proj.cout_total_investissements
    sch = [Ligne(a.libelle, [a.montant, a.montant / total if total else None]) for a in proj.entree.apports]
    sch += [Ligne(f"{p.libelle_type} — {p.libelle}", [p.montant, p.montant / total if total else None])
            for p in proj.entree.prets if p.inclure_plan_financement]
    if sch:
        sch.append(Ligne("TOTAL", [sum(l.valeurs[0] for l in sch),
                                   sum(l.valeurs[0] for l in sch) / total if total else None], "total"))
        R.tableau(["Schéma de financement", "Montant (FCFA)", "Part"], sch, largeur_libelle=R.largeur - 7.0,
                  styles_colonnes=["", "pct"], taille=9)


def _hypotheses_ca(R: Redacteur, proj: Projection):
    e = proj.entree
    lignes = []
    for prod in e.produits:
        if prod.montants is not None:
            lignes.append(Ligne(f"{prod.libelle} — chiffre d'affaires", prod.calculer(e.n)))
            continue
        q = [prod.quantites[a] if not isinstance(prod.quantites, (int, float)) else prod.quantites for a in range(e.n)]
        pu = [prod.prix_unitaire] * e.n if isinstance(prod.prix_unitaire, (int, float)) else list(prod.prix_unitaire)
        unite = f" ({prod.unite})" if prod.unite else ""
        lignes += [Ligne(f"{prod.libelle} — quantités{unite}", q, "detail"),
                   Ligne(f"{prod.libelle} — prix unitaire", pu, "detail"),
                   Ligne(f"{prod.libelle} — chiffre d'affaires", prod.calculer(e.n), "sous_total")]
    lignes.append(Ligne("Chiffre d'affaires total", proj.compte.chiffre_affaires, "total"))
    R.tableau(["Rubriques"] + list(e.annees), lignes, largeur_libelle=5.2, taille=7.5)


def _previsions(R: Redacteur, d: DossierBEFA, proj: Projection, coms, nom: str, niveau_base: int = 2):
    ce = proj.compte
    R.titre("Hypothèses de chiffre d'affaires", niveau_base + 1)
    _hypotheses_ca(R, proj)
    R.titre("Compte d'exploitation prévisionnel", niveau_base + 1)
    R.tableau(["Rubriques"] + ce.annees, ce.lignes(), largeur_libelle=4.6, taille=7.5,
              titre=f"{nom} : compte d'exploitation prévisionnel (FCFA)")
    _commentaires(R, coms, "compte_exploitation")
    R.figure(graphiques.courbes_annuelles(ce.annees, [("Chiffre d'affaires", ce.chiffre_affaires),
                                                       ("Résultat net", ce.resultat_net)]),
             "évolution du chiffre d'affaires et du résultat net (millions FCFA)", 14)
    R.titre("Plan de financement", niveau_base + 1)
    R.tableau(["Rubriques"] + proj.plan.colonnes, proj.plan.lignes(), largeur_libelle=4.0, taille=7)
    _commentaires(R, coms, "plan_financement")
    R.titre("Ratio de couverture de la dette", niveau_base + 1)
    R.tableau(["Rubrique"] + proj.couverture.annees, proj.couverture.lignes(), largeur_libelle=4.6, taille=7.5)
    if any(v is not None for v in proj.couverture.dscr):
        R.tableau(["Service de la dette"] + proj.couverture.annees, proj.couverture.lignes_service_dette(),
                  largeur_libelle=4.6, taille=7.5)
    _commentaires(R, coms, "couverture_dette")
    if proj.echeanciers:
        R.titre("Échéancier annuel des emprunts", niveau_base + 1)
        for nom_pret, agg in proj.echeanciers.items():
            lignes = [Ligne("Intérêts", agg["interets"]), Ligne("Capital remboursé", agg["amortissements"]),
                      Ligne("Capital restant dû en fin d'année", agg["capital_fin"], "sous_total")]
            R.tableau([nom_pret.strip()] + ce.annees, lignes, largeur_libelle=4.6, taille=7.5)
            if agg["hors_horizon"] > 0.5:
                R.note(f"Capital restant à rembourser au-delà de la période du plan : {C.entier(agg['hors_horizon'])} FCFA.")


def _sources_remboursement(R: Redacteur, proj: Projection, nom: str):
    ce = proj.compte
    R.titre("Source principale", 3)
    R.paragraphe(f"Le remboursement des concours sollicités sera assuré par la capacité d'autofinancement générée "
                 f"par l'exploitation de {nom}, qui passe de {C.montant(ce.caf[0])} {C.en(ce.annees[0])} à "
                 f"{C.montant(ce.caf[-1])} {C.en(ce.annees[-1])}.")
    garanties = [p.garantie for p in proj.entree.prets if p.garantie]
    R.titre("Sources secondaires", 3)
    R.paragraphe("En complément : les apports en compte courant des associés et la domiciliation de l'ensemble des "
                 "recettes de l'entreprise dans les livres de la banque" +
                 (f", ainsi que les garanties proposées ({'; '.join(garanties)})." if garanties else "."))


def _trends(R: Redacteur, d: DossierBEFA, donnees: Optional[trends.DonneesTrends]):
    if not donnees or not donnees.evolution:
        return
    R.titre("Tendances de recherche en ligne (Google Trends)", 3)
    ev = donnees.evolution
    analyses = [trends.analyser_serie(s) for s in ev.series]
    R.figure(graphiques.evolution_interet([(s.terme, s.points) for s in ev.series], ev.granularite),
             "intérêt de recherche pour " + ", ".join(f"« {s.terme} »" for s in ev.series) +
             (f" en {ev.zone}" if ev.zone else ""), 15)
    principale = next((a for a in analyses if a.saisonnalite), None)
    if principale:
        R.figure(graphiques.saisonnalite(principale.saisonnalite),
                 f"saisonnalité de la recherche « {principale.terme} » (100 = moyenne annuelle)", 14)
    if donnees.regions and donnees.regions[0].regions:
        top = donnees.regions[0].regions[:10]
        R.figure(graphiques.barres_horizontales([n for n, _ in top], [v for _, v in top]),
                 f"intérêt par région pour « {donnees.regions[0].terme} » (100 = région la plus intéressée)", 13)
    R.note(trends.source(donnees) + " " + trends.AVERTISSEMENT)


def _bibliographie(R: Redacteur, red: Optional[Redaction]):
    if not red or not red.sources:
        return
    R.titre("Sources", 1)
    R.paragraphe("Statistiques et faits sectoriels cités dans ce document (recherche documentaire réalisée par "
                 "BEFA Consulting) :", taille=9.5)
    for s in red.sources:
        n = R.renvois[s.id]
        p = R.doc.add_paragraph()
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.left_indent = Cm(0.8)
        p.paragraph_format.first_line_indent = Cm(-0.8)
        r = p.add_run(f"[{n}]  ")
        r.bold, r.font.size, r.font.color.rgb = True, Pt(8.5), ORANGE_FONCE
        r = p.add_run(f"{s.titre} — ")
        r.font.size = Pt(8.5)
        r = p.add_run(s.url)
        r.font.size, r.font.color.rgb = Pt(8), GRIS


def _bilan_ouverture(R: Redacteur, d: DossierBEFA, coms):
    h = d.analyse_historique
    if h is None:
        R.paragraphe("Entreprise nouvellement créée : le bilan d'ouverture est joint en annexe du présent dossier.")
        return
    R.tableau(["Postes", h.bilan_n.exercice], [Ligne(l.libelle, l.valeurs[:1], l.style) for l in h.lignes_bilan()],
              largeur_libelle=R.largeur - 4.5, taille=8, titre=f"Bilan d'ouverture au {h.bilan_n.exercice} (FCFA)")
    _commentaires(R, coms, "analyse_bilan")


def construire_nouveau_projet(R: Redacteur, d: DossierBEFA, red, proj, donnees_trends, coms):
    e, p = d.entreprise, d.projet
    nom = e.sigle or e.raison_sociale
    _section_ia(R, red, "introduction", "INTRODUCTION", 1)

    R.titre(f"PARTIE 1 : PRÉSENTATION DE {e.raison_sociale.upper()}", 1)
    R.titre("I. Présentation de l'entreprise", 2)
    _identite(R, d)
    R.titre("Géographie de l'actionnariat", 3)
    _actionnariat(R, d)
    R.titre("II. Management", 2)
    _management(R, d)
    R.titre("III. Produits", 2)
    _produits(R, d)
    R.titre("IV. Les clients", 2)
    _partenaires(R, d.clients, "principaux clients et modes de règlement")
    R.titre("V. Les fournisseurs", 2)
    _partenaires(R, d.fournisseurs, "principaux fournisseurs")
    R.titre("VI. Pool bancaire", 2)
    R.paragraphe(d.pool_bancaire) if d.pool_bancaire else R.a_completer("relations bancaires")

    R.titre("PARTIE 2 : PRÉSENTATION DE LA REQUÊTE", 1)
    R.titre("1. Concours sollicités et caractéristiques des prêts", 2)
    if proj and proj.entree.prets:
        _prets(R, proj, coms)
    else:
        R.a_completer("concours sollicités")
    R.titre("2. Définition du projet", 2)
    R.paragraphe(p.objet or p.intitule)
    if p.description:
        R.paragraphe(p.description)

    R.titre("PARTIE 3 : ENVIRONNEMENT ET JUSTIFICATION DU PROJET", 1)
    _section_ia(R, red, "cadre_reglementaire", "1. Cadre institutionnel et réglementaire")
    _section_ia(R, red, "justification", "2. Justification du projet")

    R.titre("PARTIE 4 : LE PROJET", 1)
    R.titre("1. Objet et objectifs du projet", 2)
    if p.objet:
        R.paragraphe(p.objet)
    if p.objectifs:
        R.paragraphe("Objectifs spécifiques :")
        R.puces(p.objectifs)
    R.titre("2. Localisation et description technique", 2)
    if p.localisation:
        R.paragraphe(f"Localisation : {p.localisation}.")
    R.paragraphe(p.description) if p.description else R.a_completer("description technique du projet")
    if p.phases:
        R.titre("Planning de réalisation", 3)
        R.puces(p.phases)
    _section_ia(R, red, "marche", "3. Analyse du marché")
    _trends(R, d, donnees_trends)

    R.titre("PARTIE 5 : ORGANISATION ET GESTION DU PROJET", 1)
    _section_ia(R, red, "execution", "1. Exécution du projet")
    R.titre("2. Gestion du projet", 2)
    _management(R, d)

    R.titre("PARTIE 6 : SITUATION FINANCIÈRE DE LA SOCIÉTÉ", 1)
    _bilan_ouverture(R, d, coms)

    R.titre("PARTIE 7 : ANALYSES PRÉVISIONNELLES DU PROJET", 1)
    if proj is None:
        R.a_completer("hypothèses financières (investissements, chiffre d'affaires, charges, financements)")
    else:
        R.titre("I. Coûts du projet et besoins de financement", 2)
        _investissements(R, proj)
        R.titre("II. États financiers prévisionnels", 2)
        _previsions(R, d, proj, coms, nom)
        R.titre("III. Sources de remboursement", 2)
        _sources_remboursement(R, proj, nom)
    _section_ia(R, red, "strategie", "IV. Vision stratégique et plan de progrès")

    R.titre("PARTIE 8 : RISQUES, ANALYSE SWOT ET PERSPECTIVES", 1)
    _section_ia(R, red, "risques", "I. Risques du projet")
    R.titre("II. Analyse SWOT", 2)
    _swot(R, red)
    _section_ia(R, red, "perspectives", "III. Perspectives de développement")
    _section_ia(R, red, "impact", "IV. Impact environnemental et social")

    _section_ia(R, red, "conclusion", "CONCLUSION GÉNÉRALE", 1)


def construire_entreprise_existante(R: Redacteur, d: DossierBEFA, red, proj, donnees_trends, coms):
    e, p = d.entreprise, d.projet
    nom = e.sigle or e.raison_sociale
    R.titre("PREMIÈRE PARTIE : PRÉSENTATION", 1)
    R.titre(f"I. Présentation de {e.raison_sociale}", 2)
    _identite(R, d)
    R.titre("Géographie de l'actionnariat", 3)
    _actionnariat(R, d)
    R.titre("Le management de l'entreprise", 3)
    _management(R, d)
    _section_ia(R, red, "presentation_activite", "II. Activités et produits")
    _produits(R, d)
    R.titre("Clientèle", 3)
    _partenaires(R, d.clients, "principaux clients et modes de règlement")
    R.titre("Fournisseurs", 3)
    _partenaires(R, d.fournisseurs, "principaux fournisseurs")
    R.titre("III. Pool bancaire", 2)
    R.paragraphe(d.pool_bancaire) if d.pool_bancaire else R.a_completer("relations bancaires")

    R.titre("DEUXIÈME PARTIE : ÉTUDE FINANCIÈRE", 1)
    h = d.analyse_historique
    if h is None:
        R.a_completer("liasses fiscales des derniers exercices")
    else:
        R.titre(f"I. Analyse du bilan {h.bilan_n.exercice}", 2)
        _commentaires(R, coms, "analyse_bilan")
        R.tableau(["Postes"] + h.colonnes, h.lignes_bilan(), largeur_libelle=5.6, styles_colonnes=["", "", "", "pct"],
                  taille=7.5, titre=f"Analyse du bilan {h.bilan_n.exercice} (FCFA)")
        R.tableau(["Équilibre financier"] + h.colonnes, h.lignes_structure(), largeur_libelle=5.6,
                  styles_colonnes=["", "", "", "pct"], taille=7.5)
        R.titre("II. Analyse du compte de résultat", 2)
        _commentaires(R, coms, "analyse_compte_resultat")
        R.tableau(["Intitulés"] + h.colonnes, h.lignes_compte_resultat(), largeur_libelle=5.6,
                  styles_colonnes=["", "", "", "pct"], taille=7.5,
                  titre=f"Analyse du compte de résultat {h.cr_n.exercice} (FCFA)")

    R.titre("TROISIÈME PARTIE : PROJET ET PROJECTION", 1)
    R.titre("I. Présentation du besoin", 2)
    _section_ia(R, red, "contexte_besoin", "Contexte", 3)
    if proj is not None:
        R.titre("Budget et plan de financement du projet", 3)
        _investissements(R, proj)
        R.titre("Besoin de financement", 3)
        _prets(R, proj, coms)
    _section_ia(R, red, "impact_besoin", "Impact du besoin sur les activités de l'entreprise", 3)
    if proj is not None:
        R.titre("II. États financiers prévisionnels", 2)
        _previsions(R, d, proj, coms, nom)
        R.titre("III. Source de remboursement", 2)
        _sources_remboursement(R, proj, nom)

    R.titre("QUATRIÈME PARTIE : ÉTUDE DE MARCHÉ", 1)
    _section_ia(R, red, "contexte_secteur", "A. Contexte général du secteur")
    _section_ia(R, red, "contexte_zone", "B. Contexte spécifique de la zone d'activité")
    _section_ia(R, red, "marche", "C. Analyse du marché")
    _trends(R, d, donnees_trends)
    R.titre("D. Analyse SWOT synthétique", 2)
    _swot(R, red)
    _section_ia(R, red, "conclusion_marche", "E. Conclusion et perspectives")

    _section_ia(R, red, "conclusion", "CONCLUSION GÉNÉRALE", 1)


def construire(d: DossierBEFA, red: Optional[Redaction], proj: Optional[Projection],
               donnees_trends: Optional[trends.DonneesTrends], pages: Optional[Dict[int, int]] = None) -> Redacteur:
    R = Redacteur(pages)
    R.renvois = {s.id: i + 1 for i, s in enumerate(red.sources)} if red else {}
    coms = red.commentaires_financiers if red and red.commentaires_financiers else commentaires_financiers(d, proj)
    titre_doc = "BUSINESS PLAN" if d.variante == "nouveau_projet" else "DOSSIER DE CRÉDIT"
    R.entete_pied()
    R.page_de_garde(d, titre_doc)
    R.emplacement_sommaire()
    if d.variante == "nouveau_projet":
        construire_nouveau_projet(R, d, red, proj, donnees_trends, coms)
    else:
        construire_entreprise_existante(R, d, red, proj, donnees_trends, coms)
    _bibliographie(R, red)
    R.remplir_sommaire()
    return R


# ----------------------------------------------------------------------
# Pagination du sommaire et generation finale
# ----------------------------------------------------------------------
def _pages_sommaire(pdf: str) -> List[int]:
    from pypdf import PdfReader

    return [i for i, p in enumerate(PdfReader(pdf).pages) if MARQUE_SOMMAIRE in (p.extract_text() or "")]


def _reperer(pdf: str, sommaire: List[EntreeSommaire], exclues: List[int]) -> Dict[int, int]:
    from pypdf import PdfReader

    textes = [_normaliser(p.extract_text() or "") for p in PdfReader(pdf).pages]
    pages, curseur = {}, (max(exclues) + 1 if exclues else 1)
    for i, e in enumerate(sommaire):
        cle = _normaliser(e.titre)
        for j in range(curseur, len(textes)):
            if cle and cle in textes[j]:
                pages[i] = j + 1
                curseur = j
                break
    return pages


@dataclass
class Resultat:
    pdf: str
    docx: str
    pages: int
    sommaire_complet: bool
    alertes: List[str] = field(default_factory=list)


def generer(d: DossierBEFA, red: Optional[Redaction], donnees_trends: Optional[trends.DonneesTrends],
            dossier_sortie: str, nom_fichier: str = "business_plan") -> Resultat:
    """Genere le Business Plan (Word + PDF) avec un sommaire pagine."""
    proj = calculer_projection(d.projection) if d.projection else None
    os.makedirs(dossier_sortie, exist_ok=True)
    travail = tempfile.mkdtemp(prefix="befa_doc_")
    alertes = []

    # 1re passe : numeros provisoires, reperage des pages du sommaire et des titres
    R = construire(d, red, proj, donnees_trends, None)
    brouillon = os.path.join(travail, "passe1.docx")
    R.doc.save(brouillon)
    pdf1 = docx_vers_pdf(brouillon, travail)
    exclues = _pages_sommaire(pdf1)
    pages = _reperer(pdf1, R.sommaire, exclues)

    docx_final = os.path.join(dossier_sortie, f"{nom_fichier}.docx")
    for _ in range(2):
        R = construire(d, red, proj, donnees_trends, pages)
        R.doc.save(docx_final)
        pdf = docx_vers_pdf(docx_final, dossier_sortie)
        nouvelles = _reperer(pdf, R.sommaire, exclues)
        if nouvelles == pages:
            break
        pages = nouvelles
    manquants = [e.titre for i, e in enumerate(R.sommaire) if i not in pages]
    if manquants:
        alertes.append("Sommaire : page introuvable pour " + ", ".join(f"« {t} »" for t in manquants[:5]) + ".")
    from pypdf import PdfReader

    return Resultat(pdf, docx_final, len(PdfReader(pdf).pages), not manquants, alertes)
