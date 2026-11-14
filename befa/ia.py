"""
Acces commun a l'API Claude pour le paquet BEFA.

- Modele par defaut : claude-opus-5-5 (surchargeable via BEFA_MODEL).
- Repli automatique cote serveur ("fallbacks": "default") si le modele
  decline une requete pour raison de securite : la requete est rejouee sur
  un modele de repli dans le meme appel.
- Sortie JSON garantie par schema (output_config.format) quand un schema
  est fourni.

Toute erreur est convertie en IaError, avec un message lisible par le
consultant ; l'appelant decide si l'erreur est bloquante.
"""
import base64
import json
import os
from typing import Any, Dict, List, Optional

import anthropic

MODELE_PAR_DEFAUT = "claude-opus-5-5"
BETA_REPLI = "server-side-fallback-2026-07-01"


class IaError(RuntimeError):
    """Echec d'un appel IA (cle absente, reseau, refus, reponse invalide)."""


def modele() -> str:
    return os.environ.get("BEFA_MODEL") or MODELE_PAR_DEFAUT


def _client() -> anthropic.Anthropic:
    if not (os.environ.get("ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_AUTH_TOKEN")):
        raise IaError(
            "Aucune clé API Anthropic configurée (ANTHROPIC_API_KEY absente du fichier .env) : "
            "les fonctions IA sont indisponibles."
        )
    return anthropic.Anthropic(timeout=600.0)


def bloc_pdf(pdf: bytes, titre: Optional[str] = None) -> Dict[str, Any]:
    bloc = {
        "type": "document",
        "source": {"type": "base64", "media_type": "application/pdf",
                   "data": base64.standard_b64encode(pdf).decode("ascii")},
    }
    if titre:
        bloc["title"] = titre
    return bloc


def appeler(
    contenu: List[Dict[str, Any]],
    *,
    systeme: Optional[str] = None,
    schema: Optional[Dict[str, Any]] = None,
    outils: Optional[List[Dict[str, Any]]] = None,
    effort: str = "medium",
    max_tokens: int = 32000,
) -> Any:
    """Envoie un message utilisateur (liste de blocs) et retourne :
    - le JSON decode si `schema` est fourni ;
    - sinon la liste de tous les blocs de contenu (texte avec citations,
      appels et resultats de recherche web).
    """
    client = _client()
    params: Dict[str, Any] = dict(
        model=modele(),
        max_tokens=max_tokens,
        messages=[{"role": "user", "content": contenu}],
        output_config={"effort": effort},
        betas=[BETA_REPLI],
        fallbacks="default",
    )
    if systeme:
        params["system"] = systeme
    if schema:
        params["output_config"]["format"] = {"type": "json_schema", "schema": schema}
    if outils:
        params["tools"] = outils
    blocs: List[Any] = []
    try:
        # Avec des outils serveur (recherche web), l'API peut suspendre un long
        # tour ("pause_turn") : on le relance en renvoyant ce qui a ete produit.
        for _ in range(5):
            with client.beta.messages.stream(**params) as flux:
                reponse = flux.get_final_message()
            blocs.extend(reponse.content)
            if reponse.stop_reason != "pause_turn":
                break
            params["messages"] = params["messages"] + [{"role": "assistant", "content": reponse.content}]
    except anthropic.AuthenticationError as e:
        raise IaError("Clé API Anthropic invalide ou expirée : mettez à jour ANTHROPIC_API_KEY dans .env "
                      "puis redémarrez le serveur.") from e
    except anthropic.RateLimitError as e:
        raise IaError(f"Limite de requêtes Anthropic atteinte, réessayez dans quelques instants ({e.message}).") from e
    except anthropic.APIStatusError as e:
        if "credit balance" in str(e.message).lower():
            raise IaError("Crédit du compte API Anthropic épuisé : rechargez-le (console Anthropic, rubrique "
                          "« Plans & Billing ») puis relancez la génération.") from e
        raise IaError(f"Erreur de l'API Anthropic ({e.status_code}) : {e.message}") from e
    except anthropic.APIConnectionError as e:
        raise IaError(f"Impossible de joindre l'API Anthropic : {e}") from e

    if reponse.stop_reason == "refusal":
        raise IaError("La demande a été refusée par le modèle (filtre de sécurité).")
    if reponse.stop_reason == "max_tokens":
        raise IaError("Réponse IA tronquée (document trop volumineux) : réessayez avec moins de pages.")
    if not schema:
        return blocs  # tous les blocs, y compris ceux des tours suspendus
    texte = "".join(b.text for b in reponse.content if b.type == "text")
    try:
        return json.loads(texte)
    except json.JSONDecodeError as e:
        raise IaError(f"Réponse IA illisible (JSON invalide) : {e}") from e
