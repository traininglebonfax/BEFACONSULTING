"""
Appels IA (Lot 4) via l'API Anthropic (Claude) - deux roles distincts, cf.
cahier des charges Section2 :

1. Textes narratifs (role principal) : description de l'activite,
   justification du profil de croissance.
2. Suggestion de valeurs par defaut sectorielles (role optionnel, "a activer
   avec prudence") : pour l'instant limite au cout variable unitaire d'un
   produit quand l'utilisateur le laisse vide - l'IA propose une valeur
   plausible compte tenu du secteur et du prix de vente (equivalent a
   "suggerer un taux de marge courant pour ce secteur").

Regle d'or : l'IA ne calcule JAMAIS de grandeur financiere du modele
(CA, marge, resultat, tresorerie...) et ne modifie ni formule ni structure.
Les textes narratifs sont attaches en COMMENTAIRES de cellule (jamais dans
une cellule de saisie/calcul, voir engine.generator.attach_ai_comments).
Les couts suggeres sont ecrits dans la cellule de saisie normale du cout
variable (comme une valeur saisie par l'utilisateur - le moteur ne fait pas
de distinction cote classeur), mais restent modifiables : l'utilisateur peut
resaisir le champ correspondant dans le formulaire et regenerer l'apercu ;
un commentaire de cellule et un badge dans l'apercu web les identifient
comme suggestion IA (voir engine.generator.attach_cost_suggestion_comments
et engine.preview).
"""
import os
from typing import Dict, Optional

import anthropic

from .models import ProjetInput

DEFAULT_MODEL = os.environ.get("ANTHROPIC_MODEL", "claude-opus-5")


class AiTextError(RuntimeError):
    """Erreur lors d'un appel IA (texte narratif ou suggestion de valeur)."""


def _client() -> "anthropic.Anthropic":
    if not (os.environ.get("ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_AUTH_TOKEN")):
        raise AiTextError(
            "Aucune cle API Anthropic configuree (variable d'environnement "
            "ANTHROPIC_API_KEY absente) - les fonctionnalites IA sont desactivees."
        )
    return anthropic.Anthropic()


def _call_claude(prompt: str, model: Optional[str] = None, max_tokens: int = 1024) -> str:
    """Appelle Claude avec un prompt simple et retourne le texte de la reponse.

    Leve AiTextError sur toute erreur d'appel (cle absente, reseau, statut API).
    """
    client = _client()
    try:
        response = client.messages.create(
            model=model or DEFAULT_MODEL,
            max_tokens=max_tokens,
            messages=[{"role": "user", "content": prompt}],
        )
    except anthropic.RateLimitError as e:
        raise AiTextError(f"Limite de requetes Anthropic atteinte: {e}") from e
    except anthropic.APIStatusError as e:
        raise AiTextError(f"Erreur API Anthropic ({e.status_code}): {e.message}") from e
    except anthropic.APIConnectionError as e:
        raise AiTextError(f"Impossible de contacter l'API Anthropic: {e}") from e

    return "".join(block.text for block in response.content if block.type == "text")


def _build_prompt(projet: ProjetInput) -> str:
    produits_lines = "\n".join(
        f"- {p.nom} : prix {p.prix_unitaire:.0f} FCFA, "
        f"volume initial {p.volume_initial_mensuel:.0f} {p.unite_volume}/mois"
        for p in projet.produits
    )
    return (
        "Voici les donnees d'un projet de business plan pour une PME :\n\n"
        f"Secteur d'activite : {projet.secteur_activite}\n"
        f"Forme juridique : {projet.forme_juridique or 'non precisee'}\n"
        f"Pays / devise : {projet.pays_devise}\n"
        "Produits / services :\n"
        f"{produits_lines}\n"
        f"Taux de croissance annuel vise : {projet.taux_croissance_a2:.0%} (annee 2), "
        f"{projet.taux_croissance_a3:.0%} (annee 3)\n\n"
        "Redige deux textes courts et concrets, en francais, pour ce business plan :\n"
        "1. DESCRIPTION_ACTIVITE : 2 a 4 phrases decrivant l'activite de l'entreprise "
        "et son offre, a partir des donnees ci-dessus.\n"
        "2. JUSTIFICATION_CROISSANCE : 2 a 4 phrases justifiant le profil de croissance "
        "vise (taux ci-dessus), de facon plausible pour ce secteur.\n\n"
        "Ne calcule et n'inclus aucun chiffre d'affaires, marge, resultat ou autre "
        "grandeur financiere - contente-toi d'un texte narratif qualitatif. "
        "Reponds strictement au format suivant, sans rien ajouter avant ou apres :\n\n"
        "DESCRIPTION_ACTIVITE: <texte>\n"
        "JUSTIFICATION_CROISSANCE: <texte>\n"
    )


def _parse_response(text: str) -> dict:
    description = ""
    justification = ""
    current = None
    for line in text.splitlines():
        if line.startswith("DESCRIPTION_ACTIVITE:"):
            current = "description"
            description = line.split(":", 1)[1].strip()
        elif line.startswith("JUSTIFICATION_CROISSANCE:"):
            current = "justification"
            justification = line.split(":", 1)[1].strip()
        elif current == "description":
            description = (description + " " + line.strip()).strip()
        elif current == "justification":
            justification = (justification + " " + line.strip()).strip()
    return {
        "description_activite": description.strip(),
        "justification_croissance": justification.strip(),
    }


def generate_narrative_texts(projet: ProjetInput, model: Optional[str] = None) -> dict:
    """Appelle Claude pour generer les 2 textes narratifs du projet.

    Retourne {"description_activite": str, "justification_croissance": str}.
    Leve AiTextError si la cle API est absente ou si l'appel/la reponse est invalide.
    """
    text = _call_claude(_build_prompt(projet), model=model)
    result = _parse_response(text)
    if not result["description_activite"] or not result["justification_croissance"]:
        raise AiTextError("Reponse IA incomplete ou mal formatee (texte narratif manquant).")
    return result


def _build_cost_prompt(projet: ProjetInput, missing_indices: list) -> str:
    lignes = "\n".join(
        f"{i}. {projet.produits[i].nom} : prix de vente {projet.produits[i].prix_unitaire:.0f} FCFA/unite"
        for i in missing_indices
    )
    return (
        f"Secteur d'activite : {projet.secteur_activite}\n"
        f"Pays / devise : {projet.pays_devise}\n\n"
        "Pour chacun des produits/services suivants, dont le prix de vente est connu mais "
        "pas le cout variable unitaire, propose une valeur plausible de COUT VARIABLE "
        "UNITAIRE (matiere premiere + emballage + production, en FCFA), coherente avec un "
        "taux de marge courant pour ce secteur d'activite en Afrique de l'Ouest :\n\n"
        f"{lignes}\n\n"
        "Ne calcule aucune autre grandeur financiere (CA, resultat, marge en valeur...) - "
        "uniquement le cout variable unitaire de chaque produit, en FCFA. "
        "Reponds strictement avec une ligne par produit, au format suivant, sans rien "
        "ajouter avant ou apres :\n\n"
        "<index>: <cout_variable_unitaire_en_FCFA>\n"
    )


def _parse_cost_suggestions(text: str, missing_indices: list) -> Dict[int, float]:
    result: Dict[int, float] = {}
    valid_indices = set(missing_indices)
    for line in text.splitlines():
        line = line.strip()
        if not line or ":" not in line:
            continue
        idx_part, _, value_part = line.partition(":")
        try:
            idx = int(idx_part.strip())
        except ValueError:
            continue
        if idx not in valid_indices:
            continue
        digits = "".join(c for c in value_part if c.isdigit() or c == ".")
        if not digits:
            continue
        try:
            result[idx] = float(digits)
        except ValueError:
            continue
    return result


def suggest_variable_costs(
    projet: ProjetInput, missing_indices: list, model: Optional[str] = None
) -> Dict[int, float]:
    """Suggere un cout variable unitaire plausible (role optionnel du Lot 4,
    cahier des charges Section2) pour les produits de `missing_indices` dont le
    cout variable unitaire est absent, a partir du secteur et du prix de vente.

    Retourne {index_produit: cout_suggere}. Leve AiTextError si la cle API est
    absente, l'appel echoue, ou la reponse ne couvre pas tous les produits demandes.
    """
    if not missing_indices:
        return {}

    text = _call_claude(_build_cost_prompt(projet, missing_indices), model=model, max_tokens=512)
    result = _parse_cost_suggestions(text, missing_indices)
    if set(result.keys()) != set(missing_indices):
        raise AiTextError("Reponse IA incomplete : cout variable manquant pour au moins un produit.")
    return result
