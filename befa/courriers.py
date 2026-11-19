"""
Demandes de financement bancaires : UN courrier distinct par banque choisie.

  1. `preparer` : texte de chaque courrier, construit par le PROGRAMME a partir
     du dossier (identite, historique, projet, financement sollicite, chiffres
     cles du Business Plan) -> aucun chiffre invente ; le consultant le modifie ;
  2. `generer`  : Word + PDF de chaque courrier, sur le papier a en-tete du
     client (Word, image ou PDF importe) ou, a defaut, avec son logo et ses
     coordonnees ; archive ZIP de l'ensemble.
"""
import io
import os
import re
import zipfile
from datetime import date
from typing import Any, Dict, List, Optional

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from . import commentaires as C
from .dossier import DossierBEFA
from .pdf import docx_vers_pdf
from .projection import Projection
from .visuel import VisuelError, entete, logo_png, proportions

MOIS = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre",
        "novembre", "décembre"]
POLICE = "Arial"
GRIS = RGBColor(0x5B, 0x61, 0x68)


def date_longue(j: Optional[date] = None) -> str:
    j = j or date.today()
    return f"{j.day if j.day > 1 else '1er'} {MOIS[j.month - 1]} {j.year}"


def minuscule(t: str) -> str:
    """Premiere lettre en minuscule, sauf sigle ou nom propre en capitales (« DAT », « BNI »)."""
    t = (t or "").strip()
    return t[0].lower() + t[1:] if len(t) > 1 and t[0].isupper() and t[1].islower() else t


def _taux(t: float) -> str:
    return f"{t * 100:.2f}".rstrip("0").rstrip(".").replace(".", ",") + " %"


def signataire(d: DossierBEFA) -> Dict[str, str]:
    x = d.dirigeants[0] if d.dirigeants else None
    return {"nom": x.nom if x else "", "fonction": (x.fonction or x.qualite) if x else "La Direction"}


def bloc_destinataire(b: Dict[str, Any]) -> List[str]:
    lignes = [b.get("titre_destinataire") or "Monsieur le Directeur Général", b.get("denomination", "")]
    if b.get("agence"):
        lignes.append(b["agence"])
    for k in ("adresse_agence", "adresse", "boite_postale"):
        if b.get(k):
            lignes.append(b[k])
            if k == "adresse_agence":
                break
    if b.get("ville") and not any(b["ville"].lower() in l.lower() for l in lignes[2:]):
        lignes.append(b["ville"])
    return [l for l in lignes if l]


def _phrase_identite(d: DossierBEFA) -> str:
    e, p = d.entreprise, d.projet
    if d.startup:
        porteur = d.dirigeants[0].nom if d.dirigeants else ""
        return (f"{d.raison_sociale}{', ' + e.forme_juridique if e.forme_juridique else ''}"
                f"{' en cours de création' if not e.rccm else ''}"
                f"{', projet porté par ' + porteur if porteur else ''}, ambitionne de développer une activité dans le "
                f"secteur « {p.secteur}{' — ' + p.sous_secteur if p.sous_secteur else ''} »"
                f"{' à ' + p.localisation if p.localisation else ''}.")
    morceaux = [d.raison_sociale]
    if e.forme_juridique:
        morceaux.append(e.forme_juridique + (f" au capital de {C.entier(e.capital)} FCFA" if e.capital else ""))
    if e.rccm:
        morceaux.append(f"immatriculée au RCCM sous le numéro {e.rccm}")
    if e.siege or e.ville:
        morceaux.append(f"dont le siège social est situé à {e.siege or e.ville}")
    activite = e.activite or e.domaine_activite or p.secteur
    return ", ".join(morceaux) + (f", exerce dans le domaine suivant : {minuscule(activite).rstrip('.')}." if activite else ".")


def preparer(d: DossierBEFA, proj: Optional[Projection], banque: Dict[str, Any], ville: str = "") -> Dict[str, Any]:
    """Texte (modifiable) du courrier adresse a une banque."""
    p = proj
    prets = p.entree.prets if p else []
    montant = sum(x.montant for x in prets)
    objet = (f"Demande de financement de {C.entier(montant)} FCFA — {d.projet.intitule}" if montant
             else f"Demande d'accompagnement financier — {d.projet.intitule}")
    titre = banque.get("titre_destinataire") or "Monsieur le Directeur Général"
    corps = [f"{titre},", _phrase_identite(d)]
    h = d.analyse_historique
    if h is not None and not d.startup:
        corps.append(f"Au titre de l'exercice {h.cr_n.exercice}, l'entreprise a réalisé un chiffre d'affaires de "
                     f"{C.montant(h.cr_n.chiffre_affaires)} pour un résultat net de {C.montant(h.cr_n.resultat_net)}, "
                     f"et dispose de capitaux propres de {C.montant(h.bilan_n.capitaux_propres)}.")
    projet = d.projet.objet or d.projet.intitule
    if p is not None and p.cout_total_investissements:
        apport = (f", financé à hauteur de {C.montant(p.total_apports)} "
                  f"({C.pct(p.total_apports / p.cout_total_investissements)}) par l'apport du promoteur"
                  if p.total_apports else "")
        corps.append(f"Dans le cadre de son développement, {d.nom} projette : {minuscule(projet).rstrip('.')}. Le coût total du "
                     f"projet est estimé à {C.montant(p.cout_total_investissements)}{apport}.")
    else:
        corps.append(f"Dans le cadre de son développement, {d.nom} projette : {minuscule(projet).rstrip('.')}.")
    nom_banque = banque.get("sigle") or banque.get("denomination", "votre établissement")
    if prets:
        details = []
        for x in prets:
            differe = " et ".join(t for t in (
                f"{x.differe_total_mois} mois de différé total" if x.differe_total_mois else "",
                f"{x.differe_partiel_mois} mois de différé partiel" if x.differe_partiel_mois else "") if t)
            details.append(f"un {x.libelle_type.lower()} de {C.entier(x.montant)} FCFA au taux indicatif de "
                           f"{_taux(x.taux_annuel)} sur {x.duree_mois} mois{' dont ' + differe if differe else ''}"
                           f"{', garanti par ' + minuscule(x.garantie).rstrip('.') if x.garantie else ''}")
        corps.append(f"À cet effet, nous avons l'honneur de solliciter auprès de {nom_banque} "
                     + " ainsi qu'".join(details) + ".")
    else:
        corps.append(f"À cet effet, nous avons l'honneur de solliciter l'accompagnement financier de {nom_banque} "
                     "pour la réalisation de ce projet.")
    if p is not None:
        ce, cv = p.compte, p.couverture
        phrase = (f"Les prévisions établies dans le Business Plan font ressortir un chiffre d'affaires passant de "
                  f"{C.montant(ce.chiffre_affaires[0])} {C.en(ce.annees[0])} à {C.montant(ce.chiffre_affaires[-1])} "
                  f"{C.en(ce.annees[-1])}, et une capacité d'autofinancement de {C.montant(ce.caf[0])} "
                  f"{C.des(ce.annees[0])}")
        dscr = [r for r in cv.dscr if r is not None]
        if dscr:
            phrase += (", qui couvre le service de la dette sur toute la durée du prêt" if min(dscr) >= 1 else
                       f" ; le taux de couverture du service de la dette atteint au minimum "
                       f"{min(dscr):.2f}".replace(".", ","))
        corps.append(phrase + ".")
    corps.append(f"{d.nom} s'engage à domicilier ses recettes dans les livres de {nom_banque}. Vous trouverez ci-joint le "
                 "Business Plan détaillé du projet. Nous restons à votre disposition pour tout complément "
                 "d'information et pour une rencontre avec vos services.")
    corps.append(f"Dans l'attente d'une suite favorable, nous vous prions d'agréer, {titre}, l'expression de notre "
                 "haute considération.")
    sig = signataire(d)
    return {
        "banque_id": banque.get("id", ""), "banque": banque.get("denomination", ""),
        "lieu_date": f"{ville or d.entreprise.ville or 'Abidjan'}, le {date_longue()}",
        "destinataire": bloc_destinataire(banque),
        "attention": f"À l'attention de {banque['destinataire']}" if banque.get("destinataire") else "",
        "objet": objet, "pieces_jointes": "Business Plan du projet",
        "corps": corps, "signataire": sig["nom"], "fonction": sig["fonction"],
        "a_verifier": bool(banque.get("a_verifier")) or not (banque.get("adresse") or banque.get("adresse_agence")),
    }


# ----------------------------------------------------------------------
# Papier du client (partage avec les autres documents)
# ----------------------------------------------------------------------
def _image_arriere_plan(run, png: bytes, largeur_cm: float, hauteur_cm: float):
    """Image ancree derriere le texte, en pleine page (papier a en-tete scanne)."""
    run.add_picture(io.BytesIO(png), width=Cm(largeur_cm), height=Cm(hauteur_cm))
    inline = run._r.find(".//" + qn("wp:inline"))
    anchor = OxmlElement("wp:anchor")
    for k, v in (("distT", "0"), ("distB", "0"), ("distL", "0"), ("distR", "0"), ("simplePos", "0"),
                 ("relativeHeight", "0"), ("behindDoc", "1"), ("locked", "1"), ("layoutInCell", "1"),
                 ("allowOverlap", "1")):
        anchor.set(k, v)
    sp = OxmlElement("wp:simplePos")
    sp.set("x", "0")
    sp.set("y", "0")
    anchor.append(sp)
    for axe, rel in (("positionH", "page"), ("positionV", "page")):
        pos = OxmlElement(f"wp:{axe}")
        pos.set("relativeFrom", rel)
        off = OxmlElement("wp:posOffset")
        off.text = "0"
        pos.append(off)
        anchor.append(pos)
    for enfant in list(inline):
        if enfant.tag == qn("wp:docPr"):
            anchor.append(OxmlElement("wp:wrapNone"))
        anchor.append(enfant)
    inline.getparent().replace(inline, anchor)


def _style(doc):
    st = doc.styles["Normal"]
    st.font.name = POLICE
    st.font.size = Pt(10.5)
    rpr = st.element.get_or_add_rPr()
    rpr.get_or_add_rFonts().set(qn("w:eastAsia"), POLICE)
    st.paragraph_format.space_after = Pt(6)


def papier(d: DossierBEFA, alertes: Optional[List[str]] = None) -> Document:
    """Document vierge sur le papier du client : gabarit Word, image d'en-tete, ou logo + coordonnees."""
    alertes = alertes if alertes is not None else []
    try:
        ent = entete(d.identite_visuelle.entete)
    except VisuelError as e:
        alertes.append(f"Papier à en-tête inutilisable ({e}) : présentation standard appliquée.")
        ent = None
    if ent and ent[0] == "docx":
        try:
            doc = Document(io.BytesIO(ent[1]))
            corps = doc.element.body
            for el in list(corps):
                if el.tag != qn("w:sectPr"):
                    corps.remove(el)
            if not doc.styles["Normal"].font.name:
                # police du theme Word (ex. Calibri / Cambria) : mal rendue par LibreOffice (accents) -> Arial
                _style(doc)
                polices = doc.styles["Normal"].element.rPr.rFonts
                for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
                    if polices.get(qn(attr)) is not None:
                        del polices.attrib[qn(attr)]
            return doc
        except Exception as e:  # gabarit Word corrompu
            alertes.append(f"Papier à en-tête Word illisible ({e}) : présentation standard appliquée.")
    doc = Document()
    _style(doc)
    sec = doc.sections[0]
    sec.page_height, sec.page_width = Cm(29.7), Cm(21)
    sec.left_margin = sec.right_margin = Cm(2.2)
    sec.top_margin, sec.bottom_margin = Cm(2.0), Cm(2.0)
    sec.header_distance = sec.footer_distance = Cm(0.6)
    if ent and ent[0] == "image":
        png = ent[1]
        ratio = proportions(png)
        h = sec.header.paragraphs[0]
        if ratio > 1.2:  # page A4 complete (en-tete + pied) : image en arriere-plan
            _image_arriere_plan(h.add_run(), png, 21.0, 21.0 * ratio)
            sec.top_margin, sec.bottom_margin = Cm(4.8), Cm(3.2)
        else:  # bandeau d'en-tete : toute la largeur utile, hauteur limitee a 6 cm
            h.alignment = WD_ALIGN_PARAGRAPH.CENTER
            sec.header_distance = Cm(0.5)
            h.paragraph_format.left_indent = h.paragraph_format.right_indent = Cm(-1.4)
            largeur = min(21.0 - 1.6, 6.0 / max(ratio, 0.01))
            h.add_run().add_picture(io.BytesIO(png), width=Cm(largeur))
            sec.top_margin = Cm(0.5 + largeur * ratio + 0.8)
        return doc
    # Presentation standard : logo + coordonnees en en-tete, mentions legales en pied
    e = d.entreprise
    h = sec.header.paragraphs[0]
    try:
        logo = logo_png(d.identite_visuelle.logo)
    except VisuelError:
        logo = None
    if logo:
        h.add_run().add_picture(io.BytesIO(logo), height=Cm(1.6))
        h.add_run("   ")
    r = h.add_run(d.raison_sociale.upper())
    r.bold, r.font.size = True, Pt(13)
    adresse = e.adresse or e.siege
    ville = e.ville if e.ville and e.ville.lower() not in (adresse or "").lower() else ""
    coord = " — ".join(x for x in (adresse, ville, f"Tél. : {e.telephone}" if e.telephone else "", e.email) if x)
    if coord:
        q = sec.header.add_paragraph()
        r = q.add_run(coord)
        r.font.size, r.font.color.rgb = Pt(8.5), GRIS
    sep = sec.header.add_paragraph()
    ppr = sep._p.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr")
    b = OxmlElement("w:bottom")
    for k, v in (("w:val", "single"), ("w:sz", "8"), ("w:space", "1"), ("w:color", "F27F24")):
        b.set(qn(k), v)
    bdr.append(b)
    ppr.append(bdr)
    mentions = " — ".join(x for x in (
        f"{e.forme_juridique} au capital de {C.entier(e.capital)} FCFA" if e.forme_juridique and e.capital else e.forme_juridique,
        f"RCCM : {e.rccm}" if e.rccm else "", f"NCC : {e.ncc}" if e.ncc else "") if x)
    if mentions:
        f = sec.footer.paragraphs[0]
        f.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = f.add_run(mentions)
        r.font.size, r.font.color.rgb = Pt(7.5), GRIS
    return doc


def _para(doc, texte="", gras=False, aligne=None, apres=6, taille=None, italique=False):
    p = doc.add_paragraph()
    if aligne is not None:
        p.alignment = aligne
    p.paragraph_format.space_after = Pt(apres)
    if texte:
        from .document import insecable

        r = p.add_run(insecable(texte))
        r.bold, r.italic = gras, italique
        if taille:
            r.font.size = Pt(taille)
    return p


def construire_courrier(d: DossierBEFA, c: Dict[str, Any], alertes: List[str]) -> Document:
    doc = papier(d, alertes)
    _para(doc, c.get("lieu_date", ""), aligne=WD_ALIGN_PARAGRAPH.RIGHT, apres=14)
    for i, l in enumerate(c.get("destinataire") or []):
        p = _para(doc, l, gras=i == 1, apres=0)
        p.paragraph_format.left_indent = Cm(9.0)
    if c.get("attention"):
        p = _para(doc, c["attention"], apres=0, italique=True)
        p.paragraph_format.left_indent = Cm(9.0)
    _para(doc, apres=10)
    p = doc.add_paragraph()
    r = p.add_run("Objet : ")
    r.bold = True
    r = p.add_run(c.get("objet", ""))
    r.bold, r.underline = True, True
    if c.get("pieces_jointes"):
        p = doc.add_paragraph()
        r = p.add_run("P.J. : ")
        r.bold = True
        p.add_run(c["pieces_jointes"])
    _para(doc, apres=6)
    corps = c.get("corps") or []
    for i, t in enumerate(corps):
        p = _para(doc, t, apres=8)
        if i > 0:
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.first_line_indent = Cm(1.0)
    _para(doc, apres=18)
    for t, g in ((c.get("fonction", ""), False), (c.get("signataire", ""), True)):
        if t:
            p = _para(doc, t, gras=g, apres=0)
            p.paragraph_format.left_indent = Cm(9.5)
    return doc


def _fichier(texte: str) -> str:
    return re.sub(r"[^\w-]+", "_", texte).strip("_")[:50] or "courrier"


def generer(d: DossierBEFA, courriers: List[Dict[str, Any]], dossier_sortie: str, pdf: bool = True) -> Dict[str, Any]:
    """Word (+ PDF) de chaque courrier et archive ZIP. Retourne la liste des fichiers produits."""
    os.makedirs(dossier_sortie, exist_ok=True)
    alertes: List[str] = []
    produits = []
    noms = set()
    for c in courriers:
        if not (c.get("banque") and c.get("corps")):
            continue
        base = f"Demande_financement_{_fichier(d.nom)}_{_fichier(c.get('banque_sigle') or c['banque'])}"
        k, nom = 2, base
        while nom in noms:
            nom, k = f"{base}_{k}", k + 1
        noms.add(nom)
        doc = construire_courrier(d, c, alertes)
        chemin = os.path.join(dossier_sortie, nom + ".docx")
        doc.save(chemin)
        entree = {"banque": c["banque"], "docx": chemin, "pdf": None}
        if pdf:
            entree["pdf"] = docx_vers_pdf(chemin, dossier_sortie)
        produits.append(entree)
    if not produits:
        raise ValueError("Aucun courrier à générer : sélectionnez au moins une banque.")
    archive = os.path.join(dossier_sortie, f"Demandes_financement_{_fichier(d.nom)}.zip")
    with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as z:
        for x in produits:
            for k in ("docx", "pdf"):
                if x[k]:
                    z.write(x[k], os.path.basename(x[k]))
    return {"courriers": produits, "zip": archive, "alertes": sorted(set(alertes))}
