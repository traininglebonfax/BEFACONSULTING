"""
Graphiques statiques (PNG) inseres dans le Business Plan PDF.

Charte : orange BEFA en serie principale ; palette categorielle validee
(lisibilite daltoniens, contraste) avec le script de reference dataviz :
#E8711A, #1F6FB2, #3A8A4A, #8E4FA6, #B8860B - attribuee dans cet ordre fixe.
Traits fins, un seul axe, grille discrete, textes en encre neutre, legende
des qu'il y a au moins deux series.
"""
import io
from datetime import date
from typing import List, Optional, Sequence, Tuple

import matplotlib

matplotlib.use("Agg")
import matplotlib.dates as mdates  # noqa: E402
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib.ticker import FuncFormatter  # noqa: E402

PALETTE = ["#E8711A", "#1F6FB2", "#3A8A4A", "#8E4FA6", "#B8860B"]
ENCRE = "#1F2328"
ENCRE_SECONDAIRE = "#5B6168"
GRILLE = "#E3E5E8"
FOND = "#FFFFFF"
MOIS_COURTS = ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc."]


def _figure(largeur=6.6, hauteur=3.0):
    fig, ax = plt.subplots(figsize=(largeur, hauteur), dpi=200)
    fig.patch.set_facecolor(FOND)
    ax.set_facecolor(FOND)
    for cote in ("top", "right", "left"):
        ax.spines[cote].set_visible(False)
    ax.spines["bottom"].set_color("#B9BEC4")
    ax.tick_params(colors=ENCRE_SECONDAIRE, labelsize=7.5, length=0)
    ax.grid(axis="y", color=GRILLE, linewidth=0.6)
    ax.set_axisbelow(True)
    return fig, ax


def _png(fig) -> bytes:
    tampon = io.BytesIO()
    fig.tight_layout()
    fig.savefig(tampon, format="png", facecolor=FOND)
    plt.close(fig)
    return tampon.getvalue()


def _moyenne_mobile(valeurs: List[float], fenetre: int) -> List[Optional[float]]:
    out: List[Optional[float]] = []
    for i in range(len(valeurs)):
        if i + 1 < fenetre:
            out.append(None)
        else:
            out.append(sum(valeurs[i + 1 - fenetre:i + 1]) / fenetre)
    return out


def evolution_interet(series: Sequence[Tuple[str, List[Tuple[date, float]]]], granularite: str = "semaine") -> bytes:
    """Courbe d'interet Google Trends. Une serie : releves bruts en trait clair +
    moyenne mobile (12 semaines / 3 mois) en trait plein. Plusieurs series :
    moyennes mobiles seules, une couleur par terme, legende."""
    fig, ax = _figure()
    fenetre = 12 if granularite == "semaine" else 3 if granularite == "mois" else 90
    libelle_mm = {"semaine": "moyenne sur 12 semaines", "mois": "moyenne sur 3 mois"}.get(granularite, "moyenne mobile")
    for i, (terme, points) in enumerate(series[: len(PALETTE)]):
        couleur = PALETTE[i]
        dates = [d for d, _ in points]
        vals = [v for _, v in points]
        mm = _moyenne_mobile(vals, fenetre)
        if len(series) == 1:
            ax.plot(dates, vals, color=couleur, linewidth=0.8, alpha=0.35, label="relevés")
            ax.plot(dates, mm, color=couleur, linewidth=2, label=libelle_mm)
        else:
            ax.plot(dates, mm, color=couleur, linewidth=2, label=terme)
    ax.set_ylim(0, 105)
    ax.set_ylabel("Indice d'intérêt (100 = pic)", fontsize=7.5, color=ENCRE_SECONDAIRE)
    ax.xaxis.set_major_locator(mdates.YearLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    ax.legend(fontsize=7.5, frameon=False, loc="upper right", labelcolor=ENCRE, ncol=min(len(series), 3) or 1)
    return _png(fig)


def saisonnalite(coefficients: List[float]) -> bytes:
    """Indice saisonnier par mois (100 = moyenne annuelle)."""
    fig, ax = _figure(hauteur=2.6)
    valeurs = [c * 100 for c in coefficients]
    ax.bar(range(12), valeurs, width=0.62, color=PALETTE[0], edgecolor=FOND, linewidth=1)
    ax.axhline(100, color=ENCRE_SECONDAIRE, linewidth=0.8, linestyle=(0, (3, 3)))
    ax.text(11.45, 101.5, "moyenne", fontsize=7, color=ENCRE_SECONDAIRE, ha="right", va="bottom")
    ax.set_xticks(range(12))
    ax.set_xticklabels(MOIS_COURTS)
    ax.set_ylim(0, max(valeurs) * 1.15)
    ax.set_ylabel("Indice saisonnier (100 = moyenne)", fontsize=7.5, color=ENCRE_SECONDAIRE)
    return _png(fig)


def barres_horizontales(libelles: List[str], valeurs: List[float], unite: str = "") -> bytes:
    """Classement (ex. regions) : barres horizontales, valeur ecrite en bout de barre."""
    n = len(libelles)
    fig, ax = _figure(hauteur=max(1.6, 0.32 * n + 0.6))
    ax.grid(False)
    ax.spines["bottom"].set_visible(False)
    y = list(range(n))[::-1]
    ax.barh(y, valeurs, height=0.6, color=PALETTE[0], edgecolor=FOND, linewidth=1)
    for yi, v in zip(y, valeurs):
        ax.text(v + max(valeurs) * 0.01, yi, f"{v:.0f}{unite}", va="center", fontsize=7.5, color=ENCRE)
    ax.set_yticks(y)
    ax.set_yticklabels(libelles, fontsize=7.5, color=ENCRE)
    ax.set_xticks([])
    ax.set_xlim(0, max(valeurs) * 1.12 if valeurs else 1)
    return _png(fig)


def courbes_annuelles(annees: List[str], series: Sequence[Tuple[str, List[float]]], unite: str = "M FCFA",
                      diviseur: float = 1e6) -> bytes:
    """Evolution annuelle de grandeurs financieres de meme unite (CA, resultat net...)."""
    fig, ax = _figure(hauteur=2.8)
    x = list(range(len(annees)))
    for i, (nom, vals) in enumerate(series[: len(PALETTE)]):
        ax.plot(x, [v / diviseur for v in vals], color=PALETTE[i], linewidth=2, marker="o", markersize=4, label=nom)
    ax.set_xticks(x)
    ax.set_xticklabels(annees)
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:,.0f}".replace(",", " ")))
    ax.set_ylabel(unite, fontsize=7.5, color=ENCRE_SECONDAIRE)
    if len(series) > 1:
        ax.legend(fontsize=7.5, frameon=False, labelcolor=ENCRE)
    return _png(fig)
