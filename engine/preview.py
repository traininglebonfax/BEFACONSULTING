"""
Conversion du classeur recalcule en tableaux HTML pour l'apercu (Lot 3).

Regle du cahier des charges (§5) : l'apercu lit les VALEURS du classeur
recalcule (jamais les formules), en conservant autant que possible la
mise en forme du modele (couleurs, gras des totaux, format FCFA).

Ce module ne modifie jamais le classeur - lecture seule, data_only=True.
"""
import html as _html
import re
from typing import List, Optional

import openpyxl
from openpyxl.utils import get_column_letter

CHAR_TO_PX = 7  # approximation largeur colonne Excel -> pixels CSS


def _rgb_to_css(argb: Optional[str]) -> Optional[str]:
    if not argb or len(argb) < 6:
        return None
    rgb = argb[-6:]
    if rgb.upper() == "000000" and len(argb) == 8 and argb.upper().startswith("00"):
        return None  # remplissage "aucune couleur" (alpha=00)
    return f"#{rgb}"


def _format_value(value, number_format: str) -> str:
    if value is None:
        return ""
    fmt = number_format or ""

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if "%" in fmt:
            decimals = fmt.count("0", fmt.find(".") + 1) if "." in fmt else 0
            return f"{value * 100:,.{decimals}f}".replace(",", " ") + "%"
        if "FCFA" in fmt:
            return f"{value:,.0f}".replace(",", " ") + " FCFA"
        if fmt in ("General", "") :
            if float(value).is_integer():
                return f"{int(value):,}".replace(",", " ")
            return f"{value:,.2f}".replace(",", " ")
        if "0.00" in fmt:
            return f"{value:,.2f}".replace(",", " ")
        if value == int(value):
            return f"{int(value):,}".replace(",", " ")
        return str(value)

    return _html.escape(str(value))


def sheet_to_html(ws) -> str:
    merges = {}  # (row, col) -> (rowspan, colspan) for top-left anchors
    covered = set()
    for rng in ws.merged_cells.ranges:
        min_row, min_col, max_row, max_col = rng.min_row, rng.min_col, rng.max_row, rng.max_col
        merges[(min_row, min_col)] = (max_row - min_row + 1, max_col - min_col + 1)
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                if (r, c) != (min_row, min_col):
                    covered.add((r, c))

    max_row = ws.max_row
    max_col = ws.max_column

    col_widths = []
    for c in range(1, max_col + 1):
        letter = get_column_letter(c)
        dim = ws.column_dimensions.get(letter)
        width_px = int((dim.width if dim and dim.width else 10) * CHAR_TO_PX)
        col_widths.append(max(40, min(width_px, 420)))

    colgroup = "".join(f'<col style="width:{w}px">' for w in col_widths)

    rows_html: List[str] = []
    for r in range(1, max_row + 1):
        row_dim = ws.row_dimensions.get(r)
        if row_dim and row_dim.hidden:
            continue

        cells_html = []
        for c in range(1, max_col + 1):
            if (r, c) in covered:
                continue
            cell = ws.cell(row=r, column=c)

            span = merges.get((r, c))
            span_attr = ""
            if span:
                rowspan, colspan = span
                if rowspan > 1:
                    span_attr += f' rowspan="{rowspan}"'
                if colspan > 1:
                    span_attr += f' colspan="{colspan}"'

            style_parts = []
            fill = cell.fill
            bg = _rgb_to_css(fill.fgColor.rgb) if fill and fill.fgColor and fill.fgColor.type == "rgb" else None
            if bg:
                style_parts.append(f"background:{bg}")
            font = cell.font
            if font:
                if font.bold:
                    style_parts.append("font-weight:700")
                if font.color and font.color.type == "rgb":
                    fc = _rgb_to_css(font.color.rgb)
                    if fc and fc.upper() not in ("#000000",):
                        style_parts.append(f"color:{fc}")
            align = cell.alignment.horizontal if cell.alignment else None
            if align:
                style_parts.append(f"text-align:{align}")
            elif isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                style_parts.append("text-align:right")

            style_attr = f' style="{";".join(style_parts)}"' if style_parts else ""
            text = _format_value(cell.value, cell.number_format)
            cells_html.append(f"<td{span_attr}{style_attr}>{text}</td>")

        rows_html.append(f"<tr>{''.join(cells_html)}</tr>")

    return f'<table class="bp-sheet"><colgroup>{colgroup}</colgroup><tbody>{"".join(rows_html)}</tbody></table>'


def workbook_to_preview(xlsx_path: str) -> List[dict]:
    """Retourne [{name, html}, ...] dans l'ordre des feuilles du classeur recalcule."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    result = []
    for name in wb.sheetnames:
        ws = wb[name]
        result.append({"name": name, "html": sheet_to_html(ws)})
    return result
