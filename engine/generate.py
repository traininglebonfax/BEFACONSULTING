"""
CLI du moteur de generation de Business Plan.

Usage:
    python -m engine.generate data/exemple_projet.json --out sortie.xlsx
    python -m engine.generate data/exemple_projet.json --out sortie.xlsx --no-recalc
"""
import argparse
import json
import re
import sys

from .generator import generate_workbook
from .models import ProjetInput
from .recalc import RecalcError, recalculate_and_check
from .structure_check import check_sheet_order


def _slug(text: str) -> str:
    text = text.strip().lower()
    text = re.sub(r"[^a-z0-9]+", "_", text).strip("_")
    return text or "business_plan"


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Genere un Business Plan Excel a partir d'un JSON de projet.")
    parser.add_argument("input_json", help="Chemin du fichier JSON decrivant le projet")
    parser.add_argument("--out", default=None, help="Chemin du fichier .xlsx de sortie (defaut: derive de nom_projet)")
    parser.add_argument("--template", default=None, help="Chemin du gabarit (defaut: model/Business_Plan_Template.xlsx)")
    parser.add_argument("--no-recalc", action="store_true", help="Ne pas recalculer via LibreOffice (debug uniquement)")
    args = parser.parse_args(argv)

    with open(args.input_json, encoding="utf-8") as f:
        data = json.load(f)

    projet = ProjetInput.from_dict(data)

    out_arg = args.out or f"{_slug(projet.nom_projet or projet.secteur_activite)}.xlsx"

    try:
        out_path = generate_workbook(projet, out_arg, template_path=args.template)
    except Exception as e:
        print(f"Erreur de generation: {e}", file=sys.stderr)
        return 1

    if not args.no_recalc:
        try:
            recalculate_and_check(out_path)
        except RecalcError as e:
            print(f"Erreur de recalcul: {e}", file=sys.stderr)
            return 2

    check_sheet_order(out_path)

    print(f"Business Plan genere: {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
