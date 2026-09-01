"""
Generation des textes narratifs (Lot 4) via l'API Anthropic (Claude).

Regle d'or (cahier des charges Section2) : l'IA ne genere QUE du texte
narratif libre - description de l'activite, justification du profil de
croissance. Elle ne calcule jamais de grandeur financiere, ne modifie
aucune formule, et n'ecrit dans aucune cellule de saisie ou de calcul :
les textes produits ici sont attaches en COMMENTAIRES de cellule sur des
cellules d'en-tete deja existantes (voir engine.generator.attach_ai_comments),
jamais dans une cellule de donnees.
"""
import os
from typing import Optional

import anthropic

from .models import ProjetInput

DEFAULT_MODEL = os.environ.get("ANTHROPIC_MODEL", "claude-opus-5")


class AiTextError(RuntimeError):
    """Erreur lors de la generation des textes narratifs par l'IA."""


def _client() -> "anthropic.Anthropic":
    if not (os.environ.get("ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_AUTH_TOKEN")):
        raise AiTextError(
            "Aucune cle API Anthropic configuree (variable d'environnement "
            "ANTHROPIC_API_KEY absente) - la generation des textes narratifs est desactivee."
        )
    return anthropic.Anthropic()


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
    client = _client()
    prompt = _build_prompt(projet)

    try:
        response = client.messages.create(
            model=model or DEFAULT_MODEL,
            max_tokens=1024,
            messages=[{"role": "user", "content": prompt}],
        )
    except anthropic.RateLimitError as e:
        raise AiTextError(f"Limite de requetes Anthropic atteinte: {e}") from e
    except anthropic.APIStatusError as e:
        raise AiTextError(f"Erreur API Anthropic ({e.status_code}): {e.message}") from e
    except anthropic.APIConnectionError as e:
        raise AiTextError(f"Impossible de contacter l'API Anthropic: {e}") from e

    text = "".join(block.text for block in response.content if block.type == "text")
    result = _parse_response(text)
    if not result["description_activite"] or not result["justification_croissance"]:
        raise AiTextError("Reponse IA incomplete ou mal formatee (texte narratif manquant).")
    return result
