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
CHART_BAR_COLOR = "#2E7D32"  # vert de marque (identique au frontend, --mid-green)

# Les emoji "pastille" (statut/tendance de 9_DASHBOARD, alerte de 5_TRESORERIE)
# n'ont pas de glyphe couleur fiable une fois passes par la police du navigateur
# (rendu noir sur blanc constate). On les remplace par un glyphe texte simple
# dont la couleur est imposee en CSS - fiable sur tout navigateur/OS, contrairement
# a un glyphe emoji "couleur" que `color:` en CSS ne peut pas teinter.
_STATUS_EMOJI_REPLACEMENTS = {
    "\U0001F7E2": ("●", "#1E4D2B"),   # cercle vert - favorable
    "\U0001F7E1": ("●", "#8a5a00"),   # cercle jaune - a surveiller
    "\U0001F534": ("●", "#b91c1c"),   # cercle rouge - defavorable
    "✅️": ("✓", "#1E4D2B"), # coche verte
    "✅": ("✓", "#1E4D2B"),
    "⚠️": ("⚠", "#b91c1c"), # triangle d'alerte
    "⚠": ("⚠", "#b91c1c"),
}


def _colorize_status_emoji(text: str) -> str:
    for emoji, (glyph, color) in _STATUS_EMOJI_REPLACEMENTS.items():
        if emoji in text:
            text = text.replace(emoji, f'<span style="color:{color};font-weight:700">{glyph}</span>')
    return text


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
            text = _colorize_status_emoji(_format_value(cell.value, cell.number_format))
            cells_html.append(f"<td{span_attr}{style_attr}>{text}</td>")

        rows_html.append(f"<tr>{''.join(cells_html)}</tr>")

    return f'<table class="bp-sheet"><colgroup>{colgroup}</colgroup><tbody>{"".join(rows_html)}</tbody></table>'


def _resolve_ref_values(wb, ref: Optional[str]) -> list:
    """Resout une reference de graphique ("'Feuille'!$A$50:$A$52") en valeurs
    de cellules deja recalculees (wb charge avec data_only=True)."""
    if not ref:
        return []
    sheet_part, _, rng = ref.partition("!")
    ws = wb[sheet_part.strip("'")]
    values = []
    for row in ws[rng.replace("$", "")]:
        for cell in row:
            values.append(cell.value)
    return values


def _extract_bar_charts(wb, ws) -> List[dict]:
    """Extrait les graphiques a barres d'une feuille sous forme de donnees
    simples (titre, categories, valeurs) - jamais les formules, seulement
    les valeurs deja recalculees pointees par le graphique."""
    charts = []
    for chart in getattr(ws, "_charts", []):
        title = None
        try:
            title = chart.title.tx.rich.p[0].r[0].t
        except Exception:
            pass
        for series in chart.series:
            val_ref = series.val.numRef.f if series.val and series.val.numRef else None
            if not val_ref:
                continue
            cat_ref = series.cat.strRef.f if series.cat and series.cat.strRef else None
            charts.append({
                "title": title,
                "categories": _resolve_ref_values(wb, cat_ref),
                "values": _resolve_ref_values(wb, val_ref),
            })
    return charts


def _bar_chart_svg(chart: dict) -> str:
    values = [v or 0 for v in chart["values"]]
    categories = chart["categories"]
    if not values:
        return ""

    n = len(values)
    max_val = max(values) or 1
    width, height = 480, 220
    margin_x, margin_top, margin_bottom = 16, 30, 30
    plot_w = width - margin_x * 2
    plot_h = height - margin_top - margin_bottom
    bar_gap = 28
    bar_w = (plot_w - bar_gap * (n - 1)) / n if n else 0
    baseline_y = margin_top + plot_h

    bars = []
    for i, v in enumerate(values):
        bar_h = (v / max_val) * plot_h if max_val else 0
        x = margin_x + i * (bar_w + bar_gap)
        y = baseline_y - bar_h
        label = _html.escape(str(categories[i]) if i < len(categories) else "")
        value_label = _html.escape(f"{v:,.0f}".replace(",", " ") + " FCFA")
        bars.append(
            f'<rect x="{x:.1f}" y="{y:.1f}" width="{bar_w:.1f}" height="{max(bar_h, 0):.1f}" '
            f'rx="4" fill="{CHART_BAR_COLOR}"><title>{label} : {value_label}</title></rect>'
            f'<text x="{x + bar_w / 2:.1f}" y="{y - 8:.1f}" text-anchor="middle" '
            f'font-size="12" font-weight="700" fill="#1f2937">{value_label}</text>'
            f'<text x="{x + bar_w / 2:.1f}" y="{height - 10:.1f}" text-anchor="middle" '
            f'font-size="12" fill="#64748b">{label}</text>'
        )

    title_html = ""
    if chart.get("title"):
        title_html = f'<div class="bp-chart-title">{_html.escape(chart["title"])}</div>'

    return (
        f'<div class="bp-chart">{title_html}'
        f'<svg viewBox="0 0 {width} {height}" role="img" '
        f'aria-label="{_html.escape(chart.get("title") or "")}">'
        f'<line x1="{margin_x}" y1="{baseline_y}" x2="{width - margin_x}" y2="{baseline_y}" '
        f'stroke="#e2e8f0" stroke-width="1"/>'
        + "".join(bars) +
        "</svg></div>"
    )


def workbook_to_preview(xlsx_path: str) -> List[dict]:
    """Retourne [{name, html}, ...] dans l'ordre des feuilles du classeur recalcule."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    result = []
    for name in wb.sheetnames:
        ws = wb[name]
        charts_html = "".join(_bar_chart_svg(c) for c in _extract_bar_charts(wb, ws))
        result.append({"name": name, "html": charts_html + sheet_to_html(ws)})
    return result
