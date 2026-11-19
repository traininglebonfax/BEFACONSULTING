"""
Identite visuelle du client : logo et papier a en-tete importes (data URL).

  - `logo_png`   : logo en PNG (Pillow normalise JPEG / WebP / GIF...) ;
  - `entete`     : papier a en-tete sous forme exploitable par les courriers :
                   ("docx", octets)  -> le document Word sert de gabarit ;
                   ("image", png)    -> image pleine page (ou bandeau) ;
                   un PDF est converti en image (1re page) par LibreOffice.
"""
import base64
import io
import os
import pathlib
import re
import shutil
import subprocess
import tempfile
from typing import Optional, Tuple

_DATA_URL = re.compile(r"^data:([\w./+-]+)?(;[\w=-]+)*;base64,(.*)$", re.S)
MIME_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


class VisuelError(ValueError):
    pass


def decoder(data_url: str) -> Tuple[str, bytes]:
    if not data_url:
        raise VisuelError("Fichier absent.")
    m = _DATA_URL.match(data_url.strip())
    if not m:
        raise VisuelError("Fichier mal encodé (data URL attendue).")
    try:
        return (m.group(1) or "application/octet-stream"), base64.b64decode(m.group(3))
    except ValueError as e:
        raise VisuelError(f"Fichier mal encodé : {e}") from e


def _en_png(octets: bytes) -> bytes:
    from PIL import Image

    try:
        img = Image.open(io.BytesIO(octets))
        img.load()
    except Exception as e:
        raise VisuelError(f"Image illisible ({e}).") from e
    if img.mode not in ("RGB", "RGBA"):
        img = img.convert("RGBA")
    sortie = io.BytesIO()
    img.save(sortie, "PNG")
    return sortie.getvalue()


def logo_png(data_url: str) -> Optional[bytes]:
    if not data_url:
        return None
    mime, octets = decoder(data_url)
    if not mime.startswith("image/") and not octets[:8].startswith(b"\x89PNG"):
        raise VisuelError("Le logo doit être une image (PNG, JPEG, WebP...).")
    return _en_png(octets)


def pdf_vers_png(pdf: bytes, timeout: int = 120) -> bytes:
    """1re page d'un PDF en image PNG, via LibreOffice (Draw)."""
    from engine.recalc import RecalcError, find_soffice

    try:
        soffice = find_soffice()
    except RecalcError as e:
        raise VisuelError("LibreOffice est nécessaire pour utiliser un papier à en-tête au format PDF.") from e
    travail = tempfile.mkdtemp(prefix="befa_entete_")
    profil = tempfile.mkdtemp(prefix="befa_lo_")
    try:
        src = os.path.join(travail, "entete.pdf")
        with open(src, "wb") as f:
            f.write(pdf)
        subprocess.run([soffice, f"-env:UserInstallation={pathlib.Path(profil).as_uri()}", "--headless",
                        "--convert-to", "png", "--outdir", travail, src],
                       capture_output=True, text=True, timeout=timeout)
        png = os.path.join(travail, "entete.png")
        if not os.path.exists(png):
            raise VisuelError("Conversion du papier à en-tête PDF impossible : importez-le en image ou en Word.")
        with open(png, "rb") as f:
            return f.read()
    except subprocess.TimeoutExpired as e:
        raise VisuelError("Conversion du papier à en-tête trop longue.") from e
    finally:
        shutil.rmtree(travail, ignore_errors=True)
        shutil.rmtree(profil, ignore_errors=True)


def entete(data_url: str) -> Optional[Tuple[str, bytes]]:
    if not data_url:
        return None
    mime, octets = decoder(data_url)
    if mime == MIME_DOCX or octets[:2] == b"PK":
        return "docx", octets
    if mime == "application/pdf" or b"%PDF" in octets[:1024]:
        return "image", pdf_vers_png(octets[octets.find(b"%PDF"):])
    if mime.startswith("image/"):
        return "image", _en_png(octets)
    raise VisuelError("Papier à en-tête : formats acceptés : image (PNG, JPEG), Word (.docx) ou PDF.")


def proportions(png: bytes) -> float:
    """Hauteur / largeur de l'image."""
    from PIL import Image

    w, h = Image.open(io.BytesIO(png)).size
    return h / w if w else 1.0
