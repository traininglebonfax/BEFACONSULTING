"""
Conversion Word -> PDF par LibreOffice (deja requis par le projet).

Chaque conversion utilise un profil LibreOffice temporaire et isole : deux
conversions simultanees (deux consultants) ne se bloquent pas.
"""
import os
import pathlib
import shutil
import subprocess
import tempfile

from engine.recalc import RecalcError, find_soffice


class PdfError(RuntimeError):
    pass


def docx_vers_pdf(docx: str, dossier_sortie: str, timeout: int = 180) -> str:
    try:
        soffice = find_soffice()
    except RecalcError as e:
        raise PdfError(str(e).replace("recalcul obligatoire avant apercu/telechargement",
                                      "génération du PDF")) from e
    docx = os.path.abspath(docx)
    os.makedirs(dossier_sortie, exist_ok=True)
    profil = tempfile.mkdtemp(prefix="befa_lo_")
    try:
        r = subprocess.run(
            [soffice, f"-env:UserInstallation={pathlib.Path(profil).as_uri()}", "--headless",
             "--convert-to", "pdf:writer_pdf_Export", "--outdir", dossier_sortie, docx],
            capture_output=True, text=True, timeout=timeout,
        )
    except subprocess.TimeoutExpired as e:
        raise PdfError(f"La conversion PDF a dépassé {timeout} s.") from e
    finally:
        shutil.rmtree(profil, ignore_errors=True)
    pdf = os.path.join(dossier_sortie, os.path.splitext(os.path.basename(docx))[0] + ".pdf")
    if r.returncode != 0 or not os.path.exists(pdf):
        raise PdfError(f"Échec de la conversion PDF par LibreOffice : {r.stderr or r.stdout}")
    return pdf
