"""Verifications structurelles du classeur genere (independantes des valeurs)."""
import openpyxl

EXPECTED_SHEETS = [
    "0_GUIDE_UTILISATION",
    "1_HYPOTHESES",
    "2_PREVISION_CA",
    "3_COMPTE_RESULTAT",
    "4_INVESTISSEMENTS",
    "5_TRESORERIE",
    "6_BFR",
    "7_RENTABILITE",
    "8_SCENARIOS",
    "9_DASHBOARD",
    "10_LEXIQUE_FINANCIER",
]


def check_sheet_order(xlsx_path: str) -> None:
    wb = openpyxl.load_workbook(xlsx_path, read_only=True)
    actual = wb.sheetnames
    if actual != EXPECTED_SHEETS:
        raise AssertionError(
            f"Ordre/nombre de feuilles incorrect.\nAttendu: {EXPECTED_SHEETS}\nObtenu:  {actual}"
        )
