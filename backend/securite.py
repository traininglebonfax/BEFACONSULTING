"""
Controle d'acces a l'application (cahier des charges 10.5).

Si la variable d'environnement BEFA_MOT_DE_PASSE est definie (fichier .env),
toutes les pages et routes exigent une authentification HTTP « Basic »
(identifiant BEFA_UTILISATEUR, « befa » par defaut) : le navigateur affiche
sa fenetre de connexion. Sans mot de passe configure, l'acces reste libre
(poste local du consultant) et un avertissement est journalise.

En production, placer l'application derriere HTTPS (proxy inverse) pour que
le mot de passe et les documents ne circulent jamais en clair.
"""
import base64
import logging
import os
import secrets

from starlette.middleware.base import BaseHTTPMiddleware
from starlette.responses import Response

LIBRES = ("/api/health",)
journal = logging.getLogger("befa.securite")


class AccesProtege(BaseHTTPMiddleware):
    async def dispatch(self, request, call_next):
        mot_de_passe = os.environ.get("BEFA_MOT_DE_PASSE", "")
        if not mot_de_passe or request.url.path in LIBRES:
            return await call_next(request)
        utilisateur = os.environ.get("BEFA_UTILISATEUR", "befa")
        entete = request.headers.get("authorization", "")
        if entete.lower().startswith("basic "):
            try:
                u, _, m = base64.b64decode(entete[6:]).decode("utf-8").partition(":")
            except (ValueError, UnicodeDecodeError):
                u, m = "", ""
            if secrets.compare_digest(u, utilisateur) and secrets.compare_digest(m, mot_de_passe):
                reponse = await call_next(request)
                reponse.headers.setdefault("Cache-Control", "no-store")
                return reponse
        return Response("Authentification requise.", status_code=401,
                        headers={"WWW-Authenticate": 'Basic realm="BEFA Consulting", charset="UTF-8"'})


class PagesFraiches(BaseHTTPMiddleware):
    """Les pages HTML ne sont pas mises en cache : apres une mise a jour, le navigateur charge la nouvelle version."""

    async def dispatch(self, request, call_next):
        reponse = await call_next(request)
        if request.url.path == "/" or request.url.path.endswith(".html"):
            reponse.headers["Cache-Control"] = "no-cache"
        return reponse


def installer(app):
    app.add_middleware(PagesFraiches)
    app.add_middleware(AccesProtege)
    if not os.environ.get("BEFA_MOT_DE_PASSE"):
        journal.warning("BEFA_MOT_DE_PASSE non défini : application accessible sans authentification.")
