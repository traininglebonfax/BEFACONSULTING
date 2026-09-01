"""
Recalcul du classeur genere via LibreOffice headless, puis verification
qu'aucune erreur de formule ne subsiste (#REF!, #DIV/0!, #VALUE!, etc.).

Le cahier des charges impose ce recalcul avant tout apercu ou telechargement :
les valeurs affichees doivent correspondre a des formules reellement evaluees.
"""
import os
import platform
import shutil
import subprocess
import tempfile

import openpyxl

FORMULA_ERROR_PREFIXES = ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#NUM!", "#N/A", "#NULL!")


class RecalcError(RuntimeError):
    pass


def find_soffice() -> str:
    candidates = [
        shutil.which("soffice"),
        shutil.which("soffice.com"),
        r"C:\Program Files\LibreOffice\program\soffice.com",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        "/usr/bin/soffice",
        "/opt/libreoffice/program/soffice",
    ]
    for c in candidates:
        if c and os.path.exists(c):
            return c
    raise RecalcError(
        "LibreOffice (soffice) introuvable. Il est requis pour le recalcul "
        "obligatoire avant apercu/telechargement (cf. cahier des charges §3)."
    )


def recalculate(xlsx_path: str, timeout: int = 120) -> str:
    """Recalcule xlsx_path via LibreOffice headless (in place).

    Retourne le chemin du fichier recalcule (identique a xlsx_path).
    """
    soffice = find_soffice()
    xlsx_path = os.path.abspath(xlsx_path)
    out_dir = tempfile.mkdtemp(prefix="bp_recalc_")
    try:
        result = subprocess.run(
            [soffice, "--headless", "--convert-to",
             "xlsx:Calc MS Excel 2007 XML", "--outdir", out_dir, xlsx_path],
            capture_output=True, text=True, timeout=timeout,
        )
        if result.returncode != 0:
            raise RecalcError(f"Echec du recalcul LibreOffice: {result.stderr}")

        recalculated = os.path.join(out_dir, os.path.basename(xlsx_path))
        if not os.path.exists(recalculated):
            raise RecalcError(
                f"Fichier recalcule introuvable apres conversion "
                f"(stdout={result.stdout!r}, stderr={result.stderr!r})."
            )
        shutil.copyfile(recalculated, xlsx_path)
        return xlsx_path
    except subprocess.TimeoutExpired as e:
        raise RecalcError(f"Le recalcul LibreOffice a depasse le delai de {timeout}s.") from e
    finally:
        shutil.rmtree(out_dir, ignore_errors=True)


def scan_formula_errors(xlsx_path: str) -> list:
    """Retourne la liste des (feuille, cellule, valeur) en erreur de formule."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    errors = []
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and v.startswith(FORMULA_ERROR_PREFIXES):
                    errors.append((sheet, cell.coordinate, v))
    return errors


def recalculate_and_check(xlsx_path: str, timeout: int = 120) -> None:
    """Recalcule puis leve RecalcError si des erreurs de formule subsistent."""
    recalculate(xlsx_path, timeout=timeout)
    errors = scan_formula_errors(xlsx_path)
    if errors:
        details = "; ".join(f"{s}!{c}={v}" for s, c, v in errors[:10])
        raise RecalcError(
            f"{len(errors)} erreur(s) de formule detectee(s) apres recalcul, "
            f"telechargement bloque. Exemples: {details}"
        )
