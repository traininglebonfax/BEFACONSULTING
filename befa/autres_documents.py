"""
« Autres documents a generer » : contrats, actes juridiques, lettres bancaires
et documents de pilotage, pre-remplis a partir du dossier.

  1. `preparer` : l'IA redige le document (JSON impose : sections, paragraphes,
     puces) a partir des seules informations du dossier ; toute information
     manquante est laissee en champ « [A completer : ...] », jamais inventee.
     Les documents chiffres (echeancier, budget, plan de financement, plan de
     tresorerie) recoivent en plus des TABLEAUX calcules par le programme.
  2. `generer`  : Word + PDF du contenu (eventuellement modifie par le consultant).

Les documents juridiques portent la mention obligatoire : projet a faire
verifier et valider par un professionnel competent avant signature.
"""
import json
import os
import re
from typing import Any, Dict, List, Optional

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt, RGBColor

from . import commentaires as C
from .courriers import date_longue, papier
from .dossier import DossierBEFA
from .ia import appeler
from .pdf import docx_vers_pdf
from .projection import Projection

AVERTISSEMENT_JURIDIQUE = ("PROJET — Document généré automatiquement à partir des informations du dossier. Il doit être "
                           "vérifié, complété et validé par un professionnel compétent (avocat, notaire, conseil "
                           "juridique) avant toute signature ou utilisation officielle.")
AVERTISSEMENT = "Document généré automatiquement à partir des informations du dossier : à relire avant utilisation."

_S = {"type": "string"}
SCHEMA = {
    "type": "object",
    "properties": {
        "titre": _S,
        "sections": {"type": "array", "items": {
            "type": "object",
            "properties": {"titre": _S, "paragraphes": {"type": "array", "items": _S},
                           "puces": {"type": "array", "items": _S}},
            "required": ["titre", "paragraphes", "puces"], "additionalProperties": False}},
        "informations_manquantes": {"type": "array", "items": _S},
    },
    "required": ["titre", "sections", "informations_manquantes"],
    "additionalProperties": False,
}

SYSTEME = """Tu es juriste et consultant au cabinet BEFA Consulting (Abidjan). Tu rédiges des documents professionnels pour des entreprises ivoiriennes, conformes au droit OHADA et au droit ivoirien (Code du travail, Code général des impôts) lorsque c'est pertinent.

Règles impératives :
- Utilise UNIQUEMENT les informations du dossier fourni (identité, dirigeants, associés, projet, chiffres). N'invente aucun nom, numéro, date, adresse, montant ni clause chiffrée.
- Toute information nécessaire mais absente est laissée sous la forme « [À compléter : description] » et listée dans "informations_manquantes".
- Style : français juridique et administratif soutenu, précis ; articles numérotés pour les contrats et statuts (« Article 1 — Objet »...).
- Une lettre comporte : formule d'appel, corps, formule de politesse ; pas de bloc d'adresse (ajouté par le programme).
- Les tableaux chiffrés éventuels sont produits par le programme : ne les recopie pas, rédige seulement leur présentation."""


def _contexte(d: DossierBEFA, proj: Optional[Projection]) -> str:
    resume = d.resume_pour_ia()
    resume["date_du_jour"] = date_longue()
    if proj is not None:
        ce = proj.compte
        resume["chiffres_previsionnels"] = {
            "annees": ce.annees, "chiffre_affaires": ce.chiffre_affaires, "resultat_net": ce.resultat_net,
            "caf": ce.caf, "cout_total_investissements": proj.cout_total_investissements,
            "apports": proj.total_apports,
            "prets": [{"libelle": p.libelle, "type": p.libelle_type, "montant": p.montant, "taux": p.taux_annuel,
                       "duree_mois": p.duree_mois, "garantie": p.garantie} for p in proj.entree.prets],
        }
    h = d.analyse_historique
    if h is not None:
        resume["dernier_exercice"] = {"exercice": h.cr_n.exercice, "chiffre_affaires": h.cr_n.chiffre_affaires,
                                      "resultat_net": h.cr_n.resultat_net, "capitaux_propres": h.bilan_n.capitaux_propres}
    return json.dumps(resume, ensure_ascii=False, indent=1)


def tableaux_calcules(d: DossierBEFA, proj: Optional[Projection], calcul: str) -> List[Dict[str, Any]]:
    """Tableaux chiffres produits par le programme : {titre, entetes, lignes: [[libelle, v1, v2...]], styles}."""
    out: List[Dict[str, Any]] = []

    def depuis_lignes(titre, entetes, lignes):
        out.append({"titre": titre, "entetes": entetes,
                    "lignes": [[l.libelle] + list(l.valeurs) for l in lignes], "styles": [l.style for l in lignes]})

    if proj is None and calcul in ("echeancier", "budget", "plan_financement"):
        return out
    if calcul == "echeancier":
        for pr in proj.entree.prets:
            lignes = [[str(e.numero), f"Mois {e.mois}", e.capital_debut, e.interet_paye, e.amortissement, e.echeance,
                       e.capital_fin] for e in pr.echeancier()]
            out.append({"titre": f"{pr.libelle_type} — {pr.libelle} : {C.entier(pr.montant)} FCFA",
                        "entetes": ["N°", "Échéance", "Capital début", "Intérêts", "Amortissement", "Échéance totale",
                                    "Capital restant dû"], "lignes": lignes, "styles": [""] * len(lignes)})
    elif calcul == "budget":
        depuis_lignes("Compte d'exploitation prévisionnel (FCFA)", ["Rubriques"] + proj.compte.annees,
                      proj.compte.lignes())
    elif calcul == "plan_financement":
        depuis_lignes("Plan de financement (FCFA)", ["Rubriques"] + proj.plan.colonnes, proj.plan.lignes())
    elif calcul == "tresorerie":
        t = d.plan_tresorerie
        if t and t.lignes:
            avec_total = any(l.total is not None for l in t.lignes)
            out.append({"titre": t.titre or "Plan de trésorerie", "entetes": ["Rubriques"] + t.periodes +
                        (["TOTAL"] if avec_total else []),
                        "lignes": [[l.libelle] + list(l.valeurs) + ([l.total] if avec_total else []) for l in t.lignes],
                        "styles": [l.style for l in t.lignes]})
        elif proj is not None:
            depuis_lignes("Flux de trésorerie annuels (plan de financement, FCFA)", ["Rubriques"] + proj.plan.colonnes,
                          proj.plan.lignes())
    return out


def preparer(d: DossierBEFA, proj: Optional[Projection], modele: Dict[str, Any], precisions: str = "") -> Dict[str, Any]:
    consigne = (f"Document à rédiger : « {modele['titre']} » ({modele.get('categorie', '')}).\n"
                f"Contenu attendu : {modele.get('consigne', '')}\n")
    if modele.get("calcul"):
        consigne += "Un tableau chiffré calculé par le programme sera inséré après ta présentation.\n"
    if precisions.strip():
        consigne += f"Précisions du consultant : {precisions.strip()}\n"
    brut = appeler([{"type": "text", "text": "DOSSIER (JSON) :\n" + _contexte(d, proj)},
                    {"type": "text", "text": consigne}], systeme=SYSTEME, schema=SCHEMA, effort="medium",
                   max_tokens=32000)
    manquantes = list(dict.fromkeys(brut.get("informations_manquantes") or []))
    return {
        "modele_id": modele.get("id", ""), "titre": brut.get("titre") or modele["titre"],
        "juridique": bool(modele.get("juridique")),
        "avertissement": AVERTISSEMENT_JURIDIQUE if modele.get("juridique") else AVERTISSEMENT,
        "lettre": modele["titre"].lower().startswith("lettre"),
        "sections": [{"titre": s.get("titre", ""), "paragraphes": [p for p in s.get("paragraphes", []) if p.strip()],
                      "puces": [p for p in s.get("puces", []) if p.strip()]} for s in brut.get("sections", [])],
        "tableaux": tableaux_calcules(d, proj, modele.get("calcul", "")),
        "informations_manquantes": manquantes,
    }


def _nombre(v) -> str:
    if v is None or v == "":
        return ""
    if isinstance(v, (int, float)):
        return "–" if abs(v) < 0.5 else f"{v:,.0f}".replace(",", " ").replace("-", "−")
    return str(v)


def generer(d: DossierBEFA, contenu: Dict[str, Any], dossier_sortie: str, pdf: bool = True) -> Dict[str, Any]:
    from .document import HEX_GRIS_CLAIR, HEX_ORANGE, HEX_ORANGE_CLAIR, _bordures, _grille, _ombre, insecable

    os.makedirs(dossier_sortie, exist_ok=True)
    alertes: List[str] = []
    doc = papier(d, alertes)
    sec = doc.sections[-1]
    largeur = (sec.page_width - sec.left_margin - sec.right_margin) / 360000
    if contenu.get("lettre"):
        p = doc.add_paragraph(f"{d.entreprise.ville or 'Abidjan'}, le {date_longue()}")
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run((contenu.get("titre") or "Document").upper())
    r.bold, r.font.size, r.font.color.rgb = True, Pt(14), RGBColor(0xB8, 0x5A, 0x12)
    av = contenu.get("avertissement")
    if av:
        tab = doc.add_table(rows=1, cols=1)
        _bordures(tab, "B85A12")
        _grille(tab, [largeur])
        c = tab.rows[0].cells[0]
        _ombre(c, "FFF4E0")
        r = c.paragraphs[0].add_run(av)
        r.italic, r.font.size, r.bold = True, Pt(8.5), bool(contenu.get("juridique"))
        doc.add_paragraph()
    for s in contenu.get("sections", []):
        if s.get("titre"):
            p = doc.add_paragraph()
            p.paragraph_format.space_before, p.paragraph_format.keep_with_next = Pt(8), True
            r = p.add_run(s["titre"])
            r.bold, r.font.size = True, Pt(11)
        for x in s.get("paragraphes", []):
            p = doc.add_paragraph(insecable(x))
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        for x in s.get("puces", []):
            doc.add_paragraph(x, style="List Bullet")
    for tb in contenu.get("tableaux", []):
        p = doc.add_paragraph()
        p.paragraph_format.keep_with_next = True
        r = p.add_run(tb.get("titre", ""))
        r.bold, r.font.size, r.font.color.rgb = True, Pt(9.5), RGBColor(0xB8, 0x5A, 0x12)
        entetes = tb["entetes"]
        n = len(entetes)
        taille = 7 if n > 8 else 8
        tab = doc.add_table(rows=1, cols=n)
        _bordures(tab)
        lib = min(4.5, largeur * 0.3) if n > 2 else largeur * 0.6
        if entetes and entetes[0] == "N°":
            lib = 1.0
        _grille(tab, [lib] + [(largeur - lib) / max(1, n - 1)] * (n - 1))
        for i, e in enumerate(entetes):
            c = tab.rows[0].cells[i]
            _ombre(c, HEX_ORANGE)
            r = c.paragraphs[0].add_run(str(e))
            r.bold, r.font.size, r.font.color.rgb = True, Pt(taille), RGBColor(0xFF, 0xFF, 0xFF)
        styles = tb.get("styles") or []
        for k, ligne in enumerate(tb["lignes"]):
            cells = tab.add_row().cells
            st = styles[k] if k < len(styles) else ""
            for i in range(n):
                v = ligne[i] if i < len(ligne) else None
                p = cells[i].paragraphs[0]
                p.paragraph_format.space_after = Pt(0)
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if i == 0 or isinstance(v, str) else WD_ALIGN_PARAGRAPH.RIGHT
                r = p.add_run(_nombre(v))
                r.font.size, r.bold = Pt(taille), st in ("total", "sous_total")
                if st == "total":
                    _ombre(cells[i], HEX_ORANGE_CLAIR)
                elif st == "sous_total":
                    _ombre(cells[i], HEX_GRIS_CLAIR)
        doc.add_paragraph()
    if contenu.get("lettre"):
        from .courriers import signataire

        sig = signataire(d)
        for txt, g in ((sig["fonction"], False), (sig["nom"], True)):
            if txt:
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Cm(9.5)
                p.paragraph_format.space_after = Pt(0)
                p.add_run(txt).bold = g
    nom = re.sub(r"[^\w-]+", "_", f"{contenu.get('titre', 'document')}_{d.nom}").strip("_")[:80]
    chemin = os.path.join(dossier_sortie, nom + ".docx")
    doc.save(chemin)
    res = {"docx": chemin, "pdf": docx_vers_pdf(chemin, dossier_sortie) if pdf else None, "alertes": alertes}
    return res
