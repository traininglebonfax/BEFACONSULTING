"""
Business Plan au format BEFA Consulting : document Word (python-docx) puis
PDF (LibreOffice). Lot 5.

Structure : plan du modele de reference « Dossier de credit 1 - O'IVOIRE »
(4 parties + conclusion generale + annexes), commun aux deux profils :
  - STARTUP (nouveau_projet)        : historique remplace par la genese du
    projet, etude financiere = bilan d'ouverture PREVISIONNEL ;
  - PME, ETI, GE (entreprise_existante) : analyse des bilans N / N-1 et
    evolution sur trois exercices.

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


def insecable(t: str) -> str:
    """Espace insecable avant « % », « : » et « FCFA » : jamais de « 100 » en fin de ligne et « % » a la suivante."""
    return re.sub(r" (%|FCFA\b)", lambda m: "\u00a0" + m.group(1), t or "")


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
        self._lettre = 0  # numerotation a), b), c)... des titres de niveau 3
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
                                                   ("Heading 3", 11, ORANGE_FONCE, 10, 4),
                                                   ("Heading 4", 10.5, ENCRE, 8, 3)):
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
        pied = st["Footer"].font  # le numero de page (champ PAGE) prend la taille du style
        pied.size, pied.color.rgb, pied.name = Pt(8), GRIS, POLICE
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
    def titre(self, texte: str, niveau: int, saut: Optional[bool] = None, sommaire: bool = True):
        if saut is None:
            saut = niveau == 1
        if saut:
            self.saut_page()
        if niveau <= 2:
            self._lettre = 0
        self.doc.add_heading(texte, level=niveau)
        if sommaire and niveau <= 4:
            self.sommaire.append(EntreeSommaire(niveau, texte))

    def sous_titre(self, texte: str):
        """Titre de niveau 3 numerote a), b), c)... comme dans le modele O'IVOIRE."""
        self.titre(f"{chr(ord('a') + self._lettre)}) {texte}", 3)
        self._lettre += 1

    def page_partie(self, texte: str):
        """Page de titre d'une partie (modele O'IVOIRE : intitule seul, centre)."""
        self.saut_page()
        for _ in range(14):
            self.doc.add_paragraph().paragraph_format.space_after = Pt(6)
        self._lettre = 0
        h = self.doc.add_heading(texte, level=1)
        h.alignment = WD_ALIGN_PARAGRAPH.CENTER
        h.paragraph_format.keep_with_next = False  # sinon le titre suit le saut de page sur la page suivante
        _bordure_paragraphe(h, "bottom", HEX_ORANGE, 12)
        self.sommaire.append(EntreeSommaire(1, texte))
        self.saut_page()

    def paysage(self, actif: bool):
        """Nouvelle section en orientation paysage (annexes a 12 colonnes) ou portrait."""
        from docx.enum.section import WD_ORIENT

        sec = self.doc.add_section(WD_SECTION.NEW_PAGE)
        sec.orientation = WD_ORIENT.LANDSCAPE if actif else WD_ORIENT.PORTRAIT
        grand, petit = Cm(29.7), Cm(21)
        sec.page_width, sec.page_height = (grand, petit) if actif else (petit, grand)
        sec.different_first_page_header_footer = False
        sec.left_margin = sec.right_margin = Cm(1.2 if actif else 1.9)
        sec.footer.is_linked_to_previous = False
        f = sec.footer.paragraphs[0]
        _bordure_paragraphe(f, "top", HEX_ORANGE, 6)
        for pos in (8.25, 16.51):  # tabulations du style « Footer » du gabarit Word : a neutraliser
            f.paragraph_format.tab_stops.add_tab_stop(Cm(pos), WD_TAB_ALIGNMENT.CLEAR)
        f.paragraph_format.tab_stops.add_tab_stop(Cm(self.largeur_section), WD_TAB_ALIGNMENT.RIGHT)
        for texte in (f"Réalisé par {CABINET}", "	Page ", None):
            r = f.add_run(texte or "")
            r.font.size, r.font.color.rgb = Pt(8), GRIS
            if texte is None:
                _champ(r, "PAGE")

    @property
    def largeur_section(self) -> float:
        s = self.doc.sections[-1]
        return (s.page_width - s.left_margin - s.right_margin) / 360000

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
            r = p.add_run(insecable(morceau))
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
                styles_colonnes: Optional[Sequence[str]] = None, taille: float = 8, titre: Optional[str] = None,
                largeur_totale: Optional[float] = None):
        """Tableau financier : 1re colonne libelle, colonnes suivantes numeriques."""
        if titre:
            p = self.paragraphe(titre, taille=9, couleur=ORANGE_FONCE)
            p.runs[0].bold = True
            p.paragraph_format.keep_with_next = True
            p.paragraph_format.space_after = Pt(3)
        ncol = len(entetes)
        if ncol >= 8 and largeur_totale is None:  # plan de financement avec annee 0 : tableau tres large
            taille, largeur_libelle = min(taille, 6.8), min(largeur_libelle, 3.3)
        t = self.doc.add_table(rows=1, cols=ncol)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        _bordures(t)
        larg_val = ((largeur_totale or self.largeur) - largeur_libelle) / max(1, ncol - 1)
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
                r = p.add_run(insecable(str(v)))
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
    def page_de_garde(self, d: DossierBEFA, titre_doc: str, logo_client: Optional[bytes] = None):
        e, p = d.entreprise, d.projet
        for _ in range(3 if logo_client else 5):
            self.doc.add_paragraph()
        if logo_client:
            q = self.doc.add_paragraph()
            q.alignment = WD_ALIGN_PARAGRAPH.CENTER
            q.paragraph_format.space_after = Pt(6)
            from .visuel import proportions

            ratio = proportions(logo_client)
            largeur = 5.0 if ratio <= 0.6 else max(2.5, 4.0 / ratio)  # hauteur <= 4 cm
            q.add_run().add_picture(io.BytesIO(logo_client), width=Cm(largeur))
        for texte, taille, couleur, gras in ((d.raison_sociale.upper(), 24, ENCRE, True),
                                             (titre_doc, 30, ORANGE, True),
                                             (p.intitule if p.intitule.strip().upper() != d.raison_sociale.strip().upper()
                                              else "", 15, GRIS, False),
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
            pf.space_after = Pt(4 if e.niveau == 1 else 1)
            pf.space_before = Pt(6 if e.niveau == 1 else 0)
            if e.niveau > 1:
                pf.left_indent = Cm(0.6 * (e.niveau - 1))
            numero = str(self.pages.get(i, "")) if self.pages is not None else MARQUE_SOMMAIRE
            r = p.add_run(f"{e.titre}\t{numero}")
            r.font.size = Pt({1: 10, 2: 9.5, 3: 9}.get(e.niveau, 8.5))
            r.italic = e.niveau == 4
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
        R.titre(f"{pr.libelle_type} — {pr.libelle}", 4, sommaire=False)
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


def _identite(R: Redacteur, d: DossierBEFA):
    e = d.entreprise
    lignes = [(k, v) for k, v in (
        ("Dénomination sociale", e.raison_sociale), ("Sigle", e.sigle), ("Forme juridique", e.forme_juridique),
        ("Capital social", f"{C.entier(e.capital)} FCFA" if e.capital else ""), ("Date de création", e.date_creation),
        ("RCCM", e.rccm), ("N° de compte contribuable", e.ncc), ("Régime fiscal", e.regime_fiscal),
        ("Centre des impôts", e.centre_impots), ("Siège social", e.siege or e.ville), ("Adresse", e.adresse),
        ("Téléphone", e.telephone), ("E-mail", e.email),
        ("Domaine d'activité", e.domaine_activite)) if v]
    if lignes:
        R.tableau_texte(None, lignes, [5.5, R.largeur - 5.5], gras_premiere_colonne=True)
    elif d.startup:
        R.note("Entreprise en cours de création : les informations d'immatriculation seront communiquées dès "
               "l'accomplissement des formalités.")
    if e.objet_social:
        R.paragraphe(f"{d.raison_sociale} a pour objet, en République de Côte d'Ivoire et dans les États parties au "
                     "Traité OHADA :")
        R.puces(e.objet_social)


def _actionnariat(R: Redacteur, d: DossierBEFA):
    if not d.actionnaires:
        R.a_completer("répartition du capital (associés, apports, parts)")
        return
    capital = d.entreprise.capital or sum(a.montant or 0 for a in d.actionnaires)
    avec_parts = any(a.nombre_parts for a in d.actionnaires)
    avec_naissance = any(a.date_naissance for a in d.actionnaires)
    entetes = ["Associé / actionnaire"] + (["Naissance"] if avec_naissance else []) + ["Nationalité"] + \
        (["Parts / actions"] if avec_parts else []) + ["Montant (FCFA)", "Part"]
    lignes = []
    for a in d.actionnaires:
        part = a.part(capital)
        lignes.append([a.nom] + ([a.date_naissance] if avec_naissance else []) + [a.nationalite] +
                      ([C.entier(a.nombre_parts) if a.nombre_parts else ""] if avec_parts else []) +
                      [C.entier(a.montant) if a.montant else "", C.pct(part) if part is not None else ""])
    total_parts = sum(a.nombre_parts or 0 for a in d.actionnaires)
    total_pct = sum(a.part(capital) or 0 for a in d.actionnaires)
    lignes.append(["TOTAL"] + ([""] if avec_naissance else []) + [""] +
                  ([C.entier(total_parts) if total_parts else ""] if avec_parts else []) +
                  [C.entier(sum(a.montant or 0 for a in d.actionnaires)), C.pct(total_pct) if total_pct else ""])
    fixe = 2.4 * (len(entetes) - 1)
    R.tableau_texte(entetes, lignes, [R.largeur - fixe] + [2.4] * (len(entetes) - 1))


def _management(R: Redacteur, d: DossierBEFA):
    if d.gouvernance:
        R.paragraphe(d.gouvernance)
    if d.dirigeants:
        if len(d.dirigeants) == 1 and not d.gouvernance:
            x = d.dirigeants[0]
            R.paragraphe(f"L'entreprise est gérée quotidiennement par {x.nom}"
                         f"{', ' + x.fonction if x.fonction else ''}"
                         f"{' (' + x.qualite + ')' if x.qualite and x.qualite != x.fonction else ''}.")
        avec_profil = any(x.profil for x in d.dirigeants)
        entetes = ["Nom et prénoms", "Fonction", "Qualité juridique"] + (["Profil / expérience"] if avec_profil else [])
        larg = [5.0, 3.8, 3.6] + ([R.largeur - 12.4] if avec_profil else [])
        if not avec_profil:
            larg = [R.largeur - 7.4, 3.8, 3.6]
        R.tableau_texte(entetes, [[x.nom, x.fonction, x.qualite] + ([x.profil] if avec_profil else [])
                                  for x in d.dirigeants], larg)
    if not d.gouvernance and not d.dirigeants:
        R.a_completer("organes de direction et équipe dirigeante")


def _partenaires(R: Redacteur, liste, type_: str):
    """Tableau des clients ou fournisseurs : colonnes affichees selon les donnees disponibles."""
    if not liste:
        R.a_completer("principaux clients et modes de règlement" if type_ == "clients" else "principaux fournisseurs")
        return
    cols = [("Nom", "nom")]
    if any(x.secteur for x in liste):
        cols.append(("Secteur" if type_ == "clients" else "Domaine", "secteur"))
    if type_ == "fournisseurs" and any(x.nature for x in liste):
        cols.append(("Biens / services fournis", "nature"))
    if any(x.montant for x in liste):
        cols.append(("CA réalisé (FCFA)" if type_ == "clients" else "Achats annuels (FCFA)", "montant"))
    if any(x.part is not None for x in liste):
        cols.append(("Part du CA" if type_ == "clients" else "Part des achats", "part"))
    for lib, attr in (("Mode de règlement", "mode_paiement"), ("Délai", "delai"), ("Localisation", "localisation")):
        if any(getattr(x, attr) for x in liste):
            cols.append((lib, attr))

    def val(x, attr):
        v = getattr(x, attr)
        if attr == "montant":
            return C.entier(v) if v else ""
        if attr == "part":
            return C.pct(v) if v is not None else ""
        return v or ""

    lignes = [[val(x, a) for _, a in cols] for x in liste]
    if any(x.montant for x in liste):
        tot = sum(x.montant or 0 for x in liste)
        tot_part = sum(x.part or 0 for x in liste)
        lignes.append(["TOTAL" if i == 0 else (C.entier(tot) if a == "montant" else
                                                (C.pct(tot_part) if a == "part" and tot_part else ""))
                       for i, (_, a) in enumerate(cols)])
    etroites = {"montant": 3.0, "part": 2.0, "delai": 2.8, "mode_paiement": 3.0, "localisation": 2.6, "secteur": 3.0,
                "nature": 3.6}
    fixe = sum(etroites.get(a, 0) for _, a in cols[1:])
    largeurs = [max(3.5, R.largeur - fixe)] + [etroites.get(a, 3.0) for _, a in cols[1:]]
    if sum(largeurs) > R.largeur:
        k = R.largeur / sum(largeurs)
        largeurs = [l * k for l in largeurs]
    R.tableau_texte([c for c, _ in cols], lignes, largeurs, taille=8.5)


def analyse_portefeuille(liste, type_: str, nom: str) -> List[str]:
    """Commentaire du programme sur la concentration du portefeuille clients / fournisseurs."""
    parts = sorted([x for x in liste if x.part is not None], key=lambda x: -x.part)
    if not parts:
        return []
    top = parts[0]
    top3 = sum(x.part for x in parts[:3])
    quoi = "chiffre d'affaires" if type_ == "clients" else "achats"
    out = [f"Le premier {'client' if type_ == 'clients' else 'fournisseur'}, {top.nom}, représente "
           f"{C.pct(top.part)} des {quoi if type_ == 'fournisseurs' else 'ventes'} de {nom}"
           + (f" et les trois premiers {C.pct(top3)}." if len(parts) >= 3 else ".")]
    if type_ == "clients":
        if top.part >= 0.30:
            out.append(f"Cette concentration traduit une dépendance commerciale significative vis-à-vis de ce premier "
                       f"client ({top.nom}), que l'entreprise devra réduire par la diversification de son portefeuille.")
        else:
            out.append("Aucun client ne dépasse 30 % du chiffre d'affaires : le portefeuille est diversifié et la "
                       "dépendance commerciale limitée.")
    elif top.part >= 0.50:
        out.append(f"L'approvisionnement dépend fortement de ce premier fournisseur ({top.nom}) : une source "
                   "alternative sécuriserait l'exploitation.")
    return out


def _produits(R: Redacteur, d: DossierBEFA):
    if not d.produits:
        return
    for x in d.produits:
        q = R.doc.add_paragraph(style="List Bullet")
        R._texte(q, x.nom + (" : " if x.description else ""), gras=True)
        if x.description:
            R._texte(q, x.description)


def _texte_ia(R: Redacteur, red: Optional[Redaction], cle: str, quoi: str):
    """Corps d'une section redigee par l'IA (sans titre)."""
    _section_ia(R, red, cle, None, quoi=quoi)


def _swot_tableau(R: Redacteur, red: Optional[Redaction]):
    """SWOT au format du modele : axes / elements identifies / impacts sur le projet."""
    if not red or not any(red.swot.values()):
        R.a_completer("analyse SWOT")
        return
    t = R.doc.add_table(rows=1, cols=3)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    _bordures(t)
    largeurs = [3.2, (R.largeur - 3.2) * 0.45, (R.largeur - 3.2) * 0.55]
    _grille(t, largeurs)
    for i, e in enumerate(("Axes stratégiques", "Éléments identifiés", "Impacts sur le projet")):
        c = t.rows[0].cells[i]
        _ombre(c, HEX_ORANGE)
        r = c.paragraphs[0].add_run(e)
        r.bold, r.font.size, r.font.color.rgb = True, Pt(9), RGBColor(0xFF, 0xFF, 0xFF)
    R._repeter_entete(t.rows[0])
    for cle, titre, fond in (("forces", "FORCES", "E3F1E6"), ("faiblesses", "FAIBLESSES", "FBE3E1"),
                             ("opportunites", "OPPORTUNITÉS", "E2ECF7"), ("menaces", "MENACES", "FCEBD9")):
        elements = red.swot.get(cle, [])
        if not elements:
            continue
        premiere = None
        for k, x in enumerate(elements):
            elem, _, impact = x.partition(" : ")
            cells = t.add_row().cells
            if k == 0:
                premiere = cells[0]
                r = cells[0].paragraphs[0].add_run(titre)
                r.bold, r.font.size = True, Pt(9)
            _ombre(cells[0], fond)
            for j, texte in ((1, elem), (2, impact)):
                p = cells[j].paragraphs[0]
                p.paragraph_format.space_after = Pt(0)
                R._texte(p, texte.strip(), taille=8.5)
        if premiere is not None and len(elements) > 1:
            premiere.merge(t.rows[len(t.rows) - 1].cells[0])
    R.doc.add_paragraph()


def _bilan_ouverture(R: Redacteur, d: DossierBEFA, coms):
    from .ouverture import MENTION

    b = d.bilan_ouverture
    if b is None and d.bilan_n is not None:  # bilan d'ouverture reel importe (liasse)
        h = d.analyse_historique
        R.tableau(["Postes", h.bilan_n.exercice], [Ligne(l.libelle, l.valeurs[:1], l.style) for l in h.lignes_bilan()],
                  largeur_libelle=R.largeur - 4.5, taille=8, titre=f"Bilan d'ouverture au {h.bilan_n.exercice} (FCFA)")
        _commentaires(R, coms, "analyse_bilan")
        return
    if b is None:
        R.a_completer("bilan d'ouverture prévisionnel (étape « Documents comptables »)")
        return
    from .historique import AnalyseHistorique, CompteResultatCondense

    h = AnalyseHistorique(b, None, CompteResultatCondense(b.exercice), None)
    lignes = [Ligne(l.libelle.replace("Emplois stables", "Actif immobilisé"), l.valeurs[:1], l.style)
              for l in h.lignes_bilan() if l.libelle != "Résultat net"]
    R.tableau(["Postes", "Montant (FCFA)"], lignes, largeur_libelle=R.largeur - 4.5, taille=8.5,
              titre="Bilan d'ouverture prévisionnel (FCFA)")
    R.note(MENTION + ("" if d.bilan_ouverture_valide else " Proposition non encore validée par le consultant."))
    for p in C.bilan_ouverture(b, d.nom):
        R.paragraphe(p)


def _historique(R: Redacteur, d: DossierBEFA, coms):
    h = d.analyse_historique
    if h is None:
        R.a_completer("bilans des trois derniers exercices (étape « Documents comptables »)")
        return
    R.titre(f"I. ANALYSE DU BILAN {h.bilan_n.exercice}", 2)
    R.paragraphe(f"Le bilan {h.bilan_n.exercice} de l'entreprise se résume dans le tableau suivant :")
    R.tableau(["Postes"] + h.colonnes, h.lignes_bilan(), largeur_libelle=5.6, styles_colonnes=["", "", "", "pct"],
              taille=7.5, titre=f"Analyse du bilan {h.bilan_n.exercice} (FCFA)")
    R.tableau(["Équilibre financier"] + h.colonnes, h.lignes_structure(), largeur_libelle=5.6,
              styles_colonnes=["", "", "", "pct"], taille=7.5)
    _commentaires(R, coms, "analyse_bilan")
    R.titre("II. ANALYSE DU COMPTE DE RÉSULTAT", 2)
    R.paragraphe(f"Le compte de résultat {h.cr_n.exercice} se résume dans le tableau ci-dessous :")
    R.tableau(["Intitulés"] + h.colonnes, h.lignes_compte_resultat(), largeur_libelle=5.6,
              styles_colonnes=["", "", "", "pct"], taille=7.5,
              titre=f"Analyse du compte de résultat {h.cr_n.exercice} (FCFA)")
    _commentaires(R, coms, "analyse_compte_resultat")
    if d.bilan_n2 is not None or d.cr_n2 is not None:
        from .historique import lignes_evolution

        cols, lignes = lignes_evolution([d.bilan_n, d.bilan_n1, d.bilan_n2], [d.cr_n, d.cr_n1, d.cr_n2])
        if lignes:
            R.titre("III. ÉVOLUTION SUR LES TROIS DERNIERS EXERCICES", 2)
            R.tableau(["Agrégats"] + cols, lignes, largeur_libelle=5.6,
                      styles_colonnes=[""] * (len(cols) - 1) + ["pct"], taille=8,
                      titre="Évolution des principaux agrégats (FCFA)")
            for p in C.evolution(cols, lignes, d.nom):
                R.paragraphe(p)


def _repartition(R: Redacteur, proj: Projection):
    """Repartition des charges de l'investissement entre le promoteur et la banque (champ `financement`)."""
    invs = proj.entree.investissements
    promoteur = [i for i in invs if (i.financement or "").lower() in ("apport", "promoteur", "fonds propres")]
    banque = [i for i in invs if (i.financement or "").lower() in ("emprunt", "banque", "credit", "crédit")]
    if not (promoteur or banque):
        R.paragraphe("L'investissement est financé globalement selon le schéma de financement présenté ci-dessus, "
                     "sans affectation de postes à une source de financement particulière.")
        return
    for titre, liste in (("L'investissement à la charge du promoteur", promoteur),
                         ("L'investissement sollicité à la banque", banque)):
        if not liste:
            continue
        R.titre(titre, 4, sommaire=False)
        lignes = [Ligne(i.libelle, [i.montant]) for i in liste]
        lignes.append(Ligne("Total", [sum(i.montant for i in liste)], "total"))
        R.tableau(["Désignation", "Montant (FCFA)"], lignes, largeur_libelle=R.largeur - 4.5, taille=9)


def _caracteristiques(R: Redacteur, d: DossierBEFA, proj: Optional[Projection], red: Optional[Redaction]):
    p = d.projet
    lignes = [("Secteur d'activité", p.secteur + (f" — {p.sous_secteur}" if p.sous_secteur else "")),
              ("Nature du besoin", p.objet or p.intitule), ("Promoteur", d.raison_sociale),
              ("Localisation", p.localisation or d.entreprise.ville)]
    if proj is not None:
        total = proj.cout_total_investissements
        lignes.append(("Coût total du projet", f"{C.entier(total)} FCFA"))
        if proj.total_apports:
            lignes.append(("Apport du promoteur", f"{C.entier(proj.total_apports)} FCFA"
                           + (f" soit {C.pct(proj.total_apports / total)}" if total else "")))
        for pr in proj.entree.prets:
            differe = " et ".join(x for x in (
                f"{pr.differe_total_mois} mois de différé total" if pr.differe_total_mois else "",
                f"{pr.differe_partiel_mois} mois de différé partiel" if pr.differe_partiel_mois else "") if x)
            taux = f"{pr.taux_annuel * 100:.2f}".rstrip("0").rstrip(".").replace(".", ",") + " %"
            lignes += [(f"Financement sollicité ({pr.libelle_type})", f"{C.entier(pr.montant)} FCFA"),
                       ("Taux", taux), ("Durée", f"{pr.duree_mois} mois" + (f" avec {differe}" if differe else ""))]
            if pr.garantie:
                lignes.append(("Garantie", pr.garantie))
    s = red.section("impact_besoin") if red else None
    if s and s.puces:
        lignes.append(("Impact attendu", "; ".join(x.split(" : ")[0].rstrip(" ;.") for x in s.puces[:4])))
    R.tableau_texte(None, [(k, v) for k, v in lignes if v], [5.0, R.largeur - 5.0], gras_premiere_colonne=True)


MODE_DECAISSEMENT = [
    "Les tirages sur la ligne de crédit se feront selon le mode opératoire suivant :",
    "Les décaissements seront faits à partir des comptes suivants :",
]


def _mode_decaissement(R: Redacteur):
    R.paragraphe(MODE_DECAISSEMENT[0])
    R.puces([MODE_DECAISSEMENT[1] + " débit du compte « crédit moyen terme », crédit du compte courant de "
             "l'entreprise ;",
             "Les règlements des charges se feront par chèque de banque ou par virement direct aux fournisseurs, sur "
             "présentation des factures pro forma ou de tout document justifiant les décaissements et des références "
             "bancaires des bénéficiaires ;",
             "Les décaissements en espèces éventuels pour diverses prestations seront communiqués à la banque "
             "préalablement aux tirages."])


def _etats_previsionnels(R: Redacteur, d: DossierBEFA, proj: Projection, coms):
    ce = proj.compte
    nom = d.nom
    R.sous_titre("Compte d'exploitation prévisionnel")
    if d.startup:
        R.paragraphe("Les états qui suivent sont des PRÉVISIONS établies à partir des hypothèses du projet ; "
                     "l'entreprise ne dispose pas encore d'historique comptable.")
    R.titre("Hypothèses de chiffre d'affaires", 4, sommaire=False)
    _hypotheses_ca(R, proj)
    R.tableau(["Rubriques"] + ce.annees, ce.lignes(), largeur_libelle=4.6, taille=7.5,
              titre=f"{nom} : compte d'exploitation prévisionnel (FCFA)")
    _commentaires(R, coms, "compte_exploitation")
    R.figure(graphiques.courbes_annuelles(ce.annees, [("Chiffre d'affaires", ce.chiffre_affaires),
                                                       ("Résultat net", ce.resultat_net)]),
             "évolution du chiffre d'affaires et du résultat net (millions FCFA)", 14)
    R.sous_titre("Plan de financement sur la durée de l'emprunt")
    R.paragraphe("Du compte d'exploitation prévisionnel ci-dessus se dégage le plan de financement suivant :")
    R.tableau(["Rubriques"] + proj.plan.colonnes, proj.plan.lignes(), largeur_libelle=4.0, taille=7)
    _commentaires(R, coms, "plan_financement")
    R.sous_titre("Ratio de couverture")
    R.tableau(["Rubrique"] + proj.couverture.annees, proj.couverture.lignes(), largeur_libelle=4.6, taille=7.5)
    if any(v is not None for v in proj.couverture.dscr):
        R.tableau(["Service de la dette"] + proj.couverture.annees, proj.couverture.lignes_service_dette(),
                  largeur_libelle=4.6, taille=7.5)
    _commentaires(R, coms, "couverture_dette")
    if proj.entree.prets:
        R.sous_titre("Mode opératoire de décaissement")
        _mode_decaissement(R)


def _source_remboursement(R: Redacteur, proj: Optional[Projection], d: DossierBEFA):
    R.titre("III. SOURCE DE REMBOURSEMENT", 2)
    R.sous_titre("Source principale")
    if proj is not None:
        ce = proj.compte
        R.paragraphe(f"Les flux générés par l'activité de {d.nom} constitueront la source principale de remboursement "
                     f"du concours sollicité : la capacité d'autofinancement prévisionnelle passe de "
                     f"{C.montant(ce.caf[0])} {C.en(ce.annees[0])} à {C.montant(ce.caf[-1])} {C.en(ce.annees[-1])}.")
    else:
        R.paragraphe("Les flux générés par l'activité constitueront la source de remboursement du concours sollicité.")
    R.sous_titre("Source secondaire")
    garanties = [p.garantie for p in proj.entree.prets if p.garantie] if proj is not None else []
    R.paragraphe("Les garanties présentées constituent les secondes sources de remboursement du concours financier "
                 "sollicité. Il s'agit notamment de :")
    R.puces([f"Domiciliation irrévocable des revenus de {d.nom} dans les livres de la banque ;"] +
            [f"{g} ;" for g in garanties] + ["Délégation d'assurance."])


def _annexes(R: Redacteur, d: DossierBEFA, proj: Optional[Projection]):
    tableaux = [t for t in (d.compte_exploitation_importe, d.plan_tresorerie) if t and t.lignes]
    echeanciers = proj.echeanciers if proj is not None else {}
    if not (tableaux or echeanciers):
        return
    R.paysage(True)
    R._lettre = 0
    R.titre("ANNEXES", 1, saut=False)
    for t in tableaux:
        R.titre(t.titre or "Tableau importé", 2)
        n = len(t.periodes)
        avec_total = any(l.total is not None for l in t.lignes)
        entetes = ["Rubriques"] + t.periodes + (["TOTAL"] if avec_total else [])
        lignes = [Ligne(l.libelle, list(l.valeurs[:n]) + [None] * (n - len(l.valeurs[:n])) +
                        ([l.total] if avec_total else []), l.style) for l in t.lignes]
        taille = 6 if len(entetes) > 12 else 7
        R.tableau(entetes, lignes, largeur_libelle=4.6, taille=taille, largeur_totale=R.largeur_section)
        R.note(f"Source : document importé par le client ({t.fichier or 'fichier'}), lu automatiquement puis vérifié "
               f"par le consultant. {'Données prévisionnelles.' if t.nature == 'previsionnel' else 'Données réalisées.'}")
    if echeanciers:
        R.titre("Échéancier annuel des emprunts", 2)
        for nom_pret, agg in echeanciers.items():
            lignes = [Ligne("Intérêts", agg["interets"]), Ligne("Capital remboursé", agg["amortissements"]),
                      Ligne("Capital restant dû en fin d'année", agg["capital_fin"], "sous_total")]
            R.tableau([nom_pret.strip()] + proj.compte.annees, lignes, largeur_libelle=5.5, taille=8,
                      largeur_totale=R.largeur_section)
            if agg["hors_horizon"] > 0.5:
                R.note(f"Capital restant à rembourser au-delà de la période du plan : {C.entier(agg['hors_horizon'])} FCFA.")


def construire_plan(R: Redacteur, d: DossierBEFA, red, proj, donnees_trends, coms):
    """Plan du modele de reference « Dossier de credit 1 - O'IVOIRE »."""
    e, p = d.entreprise, d.projet
    nom = d.nom

    # PREMIERE PARTIE : PRESENTATION
    R.page_partie("PREMIÈRE PARTIE : PRÉSENTATION")
    R.titre(f"I. PRÉSENTATION DE {d.raison_sociale.upper()}", 2)
    R.sous_titre("Genèse du projet et présentation du promoteur" if d.startup else "Historique de la création")
    _texte_ia(R, red, "historique", "historique de la création" if not d.startup else "genèse du projet")
    if e.historique:
        R.paragraphe(e.historique)
    _identite(R, d)
    R.sous_titre("Géographie de l'actionnariat")
    _actionnariat(R, d)
    R.sous_titre("Le management de l'entreprise")
    _management(R, d)
    R.titre("II. ACTIVITÉS ET PRODUITS", 2)
    R.sous_titre("Présentation de l'activité")
    _texte_ia(R, red, "presentation_activite", "présentation de l'activité")
    _produits(R, d)
    R.sous_titre("Clientèle, fournisseurs et leurs modes de règlement")
    R.titre("Clientèle", 4)
    _texte_ia(R, red, "clientele", "description de la clientèle")
    _partenaires(R, d.clients, "clients")
    for t in analyse_portefeuille(d.clients, "clients", nom):
        R.paragraphe(t)
    R.titre("Fournisseurs", 4)
    _texte_ia(R, red, "fournisseurs", "description des fournisseurs")
    _partenaires(R, d.fournisseurs, "fournisseurs")
    for t in analyse_portefeuille(d.fournisseurs, "fournisseurs", nom):
        R.paragraphe(t)
    R.titre("III. POOL BANCAIRE", 2)
    if d.pool_bancaire:
        R.paragraphe(d.pool_bancaire)
    elif d.startup:
        R.paragraphe(f"{nom}, en cours de création, n'entretient pas encore de relation bancaire établie.")
    else:
        R.a_completer("relations bancaires actuelles")
    R.paragraphe("Au regard de ses ambitions de développement et de ses besoins d'investissement, le promoteur se "
                 "déclare pleinement disposé à domicilier l'ensemble de ses flux et revenus dans les livres de "
                 "l'établissement financier qui consentira le financement sollicité.")

    # DEUXIEME PARTIE : ETUDE FINANCIERE
    R.page_partie("DEUXIÈME PARTIE : ÉTUDE FINANCIÈRE")
    if d.startup:
        R.titre("I. BILAN D'OUVERTURE PRÉVISIONNEL", 2)
        _bilan_ouverture(R, d, coms)
    else:
        _historique(R, d, coms)

    # TROISIEME PARTIE : PROJET ET PROJECTION
    R.page_partie("TROISIÈME PARTIE : PROJET ET PROJECTION")
    R.titre("I. PRÉSENTATION DU BESOIN", 2)
    R.sous_titre("Contexte")
    _texte_ia(R, red, "contexte_besoin", "contexte du besoin")
    if p.contexte_besoin:
        R.paragraphe(p.contexte_besoin)
    R.sous_titre("Budget global du projet")
    if p.description:
        R.paragraphe(p.description)
    if proj is not None:
        R.paragraphe(f"Le coût total du projet est estimé à {C.entier(proj.cout_total_investissements)} FCFA ; il est "
                     "financé selon le schéma ci-dessous.")
        _investissements(R, proj)
        R.sous_titre("Répartition des charges de l'investissement")
        _repartition(R, proj)
    else:
        R.a_completer("budget du projet (étape « Données financières »)")
    R.sous_titre("Besoin de financement")
    _texte_ia(R, red, "besoin_financement", "besoin de financement")
    if proj is not None and proj.entree.prets:
        _prets(R, proj, coms)
    R.sous_titre("Impact du besoin sur les activités de l'entreprise")
    _texte_ia(R, red, "impact_besoin", "impacts attendus du financement")
    R.sous_titre("Caractéristiques du besoin")
    _caracteristiques(R, d, proj, red)
    R.titre("II. ÉTATS FINANCIERS PRÉVISIONNELS", 2)
    if proj is None:
        R.a_completer("hypothèses financières (investissements, chiffre d'affaires, charges, financements)")
    else:
        _etats_previsionnels(R, d, proj, coms)
    _source_remboursement(R, proj, d)

    # QUATRIEME PARTIE : ETUDE DE MARCHE
    R.page_partie("QUATRIÈME PARTIE : ÉTUDE DE MARCHÉ")
    secteur = (p.sous_secteur or p.secteur or "").split(" - ")[0].upper()
    R.titre(f"A. CONTEXTE ET ANALYSE GÉNÉRALE DU SECTEUR{' : ' + secteur if secteur else ''}", 2)
    R.titre("I. CONTEXTE GÉNÉRAL DU SECTEUR", 3)
    R.titre("1. Contexte mondial du secteur", 4)
    _texte_ia(R, red, "contexte_mondial", "contexte mondial du secteur")
    R.titre("2. Contexte africain du secteur", 4)
    _texte_ia(R, red, "contexte_africain", "contexte africain du secteur")
    R.titre("3. Contexte et marché en Côte d'Ivoire", 4)
    _texte_ia(R, red, "contexte_national", "contexte ivoirien du secteur")
    zone = (p.localisation or e.ville or "zone d'activité").split(",")[0].strip().upper()
    R.titre(f"II. CONTEXTE ET ANALYSE SPÉCIFIQUE : {zone}", 3)
    _texte_ia(R, red, "contexte_zone", "contexte de la zone d'activité")
    R.titre("B. ANALYSE DU MARCHÉ", 2)
    _texte_ia(R, red, "analyse_demande", "analyse de la demande")
    _trends(R, d, donnees_trends)
    R.titre("C. ANALYSE DES PRIX", 2)
    _texte_ia(R, red, "analyse_prix", "analyse des prix")
    R.titre("D. ANALYSE DE LA CONCURRENCE", 2)
    R.titre("1) Concurrence locale ou concurrence directe", 3)
    _texte_ia(R, red, "concurrence_directe", "concurrence directe")
    R.titre("2) Concurrence régionale et nationale", 3)
    _texte_ia(R, red, "concurrence_nationale", "concurrence régionale et nationale")
    R.titre("3) Opportunité concurrentielle", 3)
    _texte_ia(R, red, "opportunite_concurrentielle", "opportunité concurrentielle")
    R.titre("E. ANALYSE SWOT SYNTHÉTIQUE", 2)
    _swot_tableau(R, red)
    R.titre("F. CONCLUSION ET PERSPECTIVES", 2)
    _texte_ia(R, red, "conclusion_marche", "conclusion de l'étude de marché")

    _section_ia(R, red, "conclusion", "CONCLUSION GÉNÉRALE", 1)


def construire(d: DossierBEFA, red: Optional[Redaction], proj: Optional[Projection],
               donnees_trends: Optional[trends.DonneesTrends], pages: Optional[Dict[int, int]] = None) -> Redacteur:
    from .visuel import VisuelError, logo_png

    R = Redacteur(pages)
    R.renvois = {s.id: i + 1 for i, s in enumerate(red.sources)} if red else {}
    coms = red.commentaires_financiers if red and red.commentaires_financiers else commentaires_financiers(d, proj)
    try:
        logo = logo_png(d.identite_visuelle.logo)
    except VisuelError:
        logo = None
    R.entete_pied()
    R.page_de_garde(d, "BUSINESS PLAN" if d.startup else "DOSSIER DE CRÉDIT", logo)
    R.emplacement_sommaire()
    construire_plan(R, d, red, proj, donnees_trends, coms)
    _bibliographie(R, red)
    _annexes(R, d, proj)
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
