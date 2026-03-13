# =============================================================================
# pdf_generator.py — Rapport PDF Executive Pitchbook — Amundi Research Grade
# Charte : Marine #002D54 | Orange #FF4F00 | BLANC #FFFFFF
# Standards visuels : fond blanc, tableaux epures, contrastes institutionnels
# Graphiques : Matplotlib (Agg) uniquement — isole de Plotly (Streamlit)
# Proportions : CHART_W = USABLE_W * 0.65 (marges elegantes)
# Ordre graphiques : 1.Type Client | 2.Region | 3.Top 10 Deals
# Filtrage fund-by-fund : pilote par fonds_perimetre depuis app.py
# =============================================================================

import io
from datetime import date
from typing import Optional

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import pandas as pd
import numpy as np

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.colors import HexColor
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    Image, PageBreak,
)
from reportlab.platypus.flowables import Flowable


# ---------------------------------------------------------------------------
# CONSTANTES MISE EN PAGE & CHARTE AMUNDI
# ---------------------------------------------------------------------------

# Couleurs institutionnelles
BLEU_MARINE  = HexColor("#002D54")
ORANGE_AMUNDI= HexColor("#FF4F00")
BLANC        = HexColor("#FFFFFF")
GRIS_CLAIR   = HexColor("#E8E8E8")
GRIS_TEXTE   = HexColor("#444444")
BLEU_MID     = HexColor("#1A5E8A")
GRIS_PIED    = HexColor("#F5F5F5")

# Hexadecimaux pour Matplotlib
HEX_MARINE   = "#002D54"
HEX_ORANGE   = "#FF4F00"
HEX_BLANC    = "#FFFFFF"
HEX_GRIS     = "#E8E8E8"
HEX_GTXT     = "#444444"
HEX_BLEU_MID = "#1A5E8A"
HEX_BLEU_PAL = "#4A8FBD"
HEX_BLEU_DEP = "#003F7A"

# Mise en page A4
PAGE_W, PAGE_H = A4
MARGIN_H       = 2.0 * cm
MARGIN_V       = 1.8 * cm
USABLE_W       = PAGE_W - 2 * MARGIN_H     # ~481 pts

# Graphiques : 65% de USABLE_W — marges elegantes (standard Amundi factsheet)
CHART_W        = USABLE_W * 0.65

# Palette Amundi — bleu marine + variations institutionnelles
PALETTE = [
    HEX_BLEU_MID,    # Bleu intermediaire — dominant
    HEX_MARINE,      # Bleu marine Amundi
    HEX_BLEU_PAL,    # Bleu pâle
    HEX_BLEU_DEP,    # Bleu profond
    "#2C7FB8",       # Bleu ocean
    "#004F8C",       # Bleu institutionnel
    "#6BAED6",       # Bleu ciel doux
    "#08519C",       # Bleu nuit
    "#9ECAE1",       # Bleu pale
    "#003060",       # Bleu nuit profond
]

# Parametres rcParams Matplotlib — charte stricte, zero bruit visuel
MPL_PARAMS = {
    "font.family":          "DejaVu Sans",
    "font.size":            8.5,
    "axes.spines.top":      False,
    "axes.spines.right":    False,
    "axes.spines.left":     True,
    "axes.spines.bottom":   True,
    "axes.edgecolor":       HEX_GRIS,
    "axes.linewidth":       0.6,
    "axes.labelcolor":      HEX_MARINE,
    "axes.labelsize":       8,
    "xtick.color":          HEX_MARINE,
    "ytick.color":          HEX_MARINE,
    "xtick.labelsize":      7.5,
    "ytick.labelsize":      7.5,
    "figure.facecolor":     HEX_BLANC,
    "axes.facecolor":       HEX_BLANC,
    "grid.color":           HEX_GRIS,
    "grid.linewidth":       0.45,
    "legend.fontsize":      7.5,
    "legend.frameon":       True,
    "legend.framealpha":    0.9,
    "legend.edgecolor":     HEX_GRIS,
}


# ---------------------------------------------------------------------------
# HELPERS GENERAUX
# ---------------------------------------------------------------------------

class ColorRect(Flowable):
    """Rectangle colore pour bandeaux de section et separateurs."""
    def __init__(self, width, height, color):
        super().__init__()
        self.width  = width
        self.height = height
        self.color  = color

    def draw(self):
        self.canv.setFillColor(self.color)
        self.canv.rect(0, 0, self.width, self.height, fill=1, stroke=0)


def anonymize_df(df: pd.DataFrame) -> pd.DataFrame:
    """Mode Comex : remplace nom_client par type_client + region."""
    df = df.copy()
    if all(c in df.columns for c in ("nom_client", "type_client", "region")):
        df["nom_client"] = df.apply(
            lambda r: f"{r['type_client']} — {r['region']}", axis=1
        )
    return df


def fmt_aum(value: float) -> str:
    """Formatage institutionnel Amundi : M EUR avec 1 decimale."""
    if value is None:
        return "—"
    try:
        v = float(value)
    except (TypeError, ValueError):
        return "—"
    if v >= 1_000_000_000:
        return f"{v / 1_000_000_000:.1f} Md EUR"
    if v >= 1_000_000:
        return f"{v / 1_000_000:.1f} M EUR"
    if v >= 1_000:
        return f"{v / 1_000:.0f} k EUR"
    return f"{v:.0f} EUR"


def _centrer_chart(elements: list, buf: io.BytesIO,
                   aspect: float, caption: str, styles: dict):
    """
    Ajoute un graphique centre dans la page avec legende et separateur.
    Padding lateral = (USABLE_W - CHART_W) / 2 pour centrage parfait.
    """
    w   = CHART_W
    h   = w * aspect
    img = Image(buf, width=w, height=h)
    pad = (USABLE_W - w) / 2

    centering = Table([[img]], colWidths=[w])
    centering.setStyle(TableStyle([
        ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
        ("LEFTPADDING",   (0, 0), (-1, -1), 0),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 0),
        ("TOPPADDING",    (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))

    # Wrapper de centrage avec padding lateral
    wrapper = Table([[Spacer(pad, 1), centering, Spacer(pad, 1)]],
                    colWidths=[pad, w, pad])
    wrapper.setStyle(TableStyle([
        ("LEFTPADDING",   (0, 0), (-1, -1), 0),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 0),
        ("TOPPADDING",    (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    elements.append(wrapper)
    elements.append(Paragraph(caption, styles["chart_caption"]))
    elements.append(Spacer(1, 8))
    elements.append(ColorRect(USABLE_W * 0.35, 1, GRIS_CLAIR))
    elements.append(Spacer(1, 14))


# ---------------------------------------------------------------------------
# GRAPHIQUES MATPLOTLIB — ISOLES DE PLOTLY (app.py)
# Chaque fonction : rcParams.update -> figure -> savefig -> close -> BytesIO
# ---------------------------------------------------------------------------

def _donut(ax, labels: list, values: list, title: str):
    """
    Donut chart Amundi : palette bleu marine, legendes compactes, charte stricte.
    Sous-routine partagee par les deux graphiques de repartition.
    """
    colors = [PALETTE[i % len(PALETTE)] for i in range(len(labels))]
    _, _, autotexts = ax.pie(
        values,
        colors=colors,
        autopct="%1.1f%%",
        startangle=90,
        pctdistance=0.74,
        wedgeprops={"width": 0.52, "edgecolor": HEX_BLANC, "linewidth": 1.8}
    )
    for at in autotexts:
        at.set_fontsize(8.5)
        at.set_color(HEX_BLANC)
        at.set_fontweight("bold")

    total = sum(values)
    ax.text(0,  0.09, fmt_aum(total),
            ha="center", va="center",
            fontsize=9, fontweight="bold", color=HEX_MARINE)
    ax.text(0, -0.19, "AUM Finance",
            ha="center", va="center", fontsize=7, color=HEX_GTXT)

    patches = [
        mpatches.Patch(color=colors[i],
                       label=f"{labels[i]}: {fmt_aum(values[i])}")
        for i in range(len(labels))
    ]
    ax.legend(
        handles=patches,
        loc="lower center",
        bbox_to_anchor=(0.5, -0.30),
        ncol=2,
        fontsize=7,
        frameon=False,
        labelcolor=HEX_MARINE
    )
    ax.set_title(title,
                 fontsize=9.5, fontweight="bold",
                 color=HEX_MARINE, pad=10)


def _chart_pie_aum_by_type(aum_by_type: dict) -> io.BytesIO:
    """Donut — AUM Funded par type de client. Matplotlib Agg uniquement."""
    plt.rcParams.update(MPL_PARAMS)
    # Dimensions : rapport CHART_W (pts) -> pouces (72 pts/inch)
    fig_w = CHART_W / 72
    fig, ax = plt.subplots(figsize=(fig_w, fig_w * 0.80))
    fig.patch.set_facecolor(HEX_BLANC)

    if not aum_by_type or sum(aum_by_type.values()) == 0:
        ax.text(0.5, 0.5, "Aucune donnee Funded",
                ha="center", va="center",
                fontsize=9, color=HEX_MARINE, transform=ax.transAxes)
        ax.axis("off")
    else:
        _donut(ax,
               list(aum_by_type.keys()),
               list(aum_by_type.values()),
               "Repartition AUM Funded — Type de Client")

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150,
                bbox_inches="tight", facecolor=HEX_BLANC)
    plt.close(fig)
    buf.seek(0)
    return buf


def _chart_pie_aum_by_region(aum_by_region: dict) -> io.BytesIO:
    """Donut — AUM Funded par region geographique. Matplotlib Agg uniquement."""
    plt.rcParams.update(MPL_PARAMS)
    fig_w = CHART_W / 72
    fig, ax = plt.subplots(figsize=(fig_w, fig_w * 0.80))
    fig.patch.set_facecolor(HEX_BLANC)

    if not aum_by_region or sum(aum_by_region.values()) == 0:
        ax.text(0.5, 0.5, "Aucune donnee Funded par region",
                ha="center", va="center",
                fontsize=9, color=HEX_MARINE, transform=ax.transAxes)
        ax.axis("off")
    else:
        _donut(ax,
               list(aum_by_region.keys()),
               list(aum_by_region.values()),
               "Repartition AUM Funded — Region Geographique")

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150,
                bbox_inches="tight", facecolor=HEX_BLANC)
    plt.close(fig)
    buf.seek(0)
    return buf


def _chart_top10_deals(top_deals: list, mode_comex: bool) -> io.BytesIO:
    """
    Bar chart horizontal — Top 10 Deals funded.
    Hauteur dynamique proportionnelle au nombre de deals.
    Matplotlib Agg uniquement.
    """
    plt.rcParams.update(MPL_PARAMS)

    deals = top_deals[:10]
    n     = len(deals)
    fig_w = CHART_W / 72
    fig_h = max(2.8, n * 0.50 + 0.9)
    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    fig.patch.set_facecolor(HEX_BLANC)
    ax.set_facecolor(HEX_BLANC)

    if not deals:
        ax.text(0.5, 0.5, "Aucun deal Funded",
                ha="center", va="center",
                fontsize=9, color=HEX_MARINE, transform=ax.transAxes)
        ax.axis("off")
    else:
        labels = [
            f"{d['type_client']} — {d['region']}" if mode_comex
            else str(d.get("nom_client", ""))
            for d in deals
        ]
        values = [float(d.get("funded_aum", 0)) for d in deals]

        # Tri decroissant
        pairs  = sorted(zip(values, labels), reverse=True)
        values = [v for v, _ in pairs]
        labels = [l for _, l in pairs]
        max_v  = max(values) if values else 1.0

        # Degradé de couleurs : premier = marine plein, suivants = degrades
        bar_colors = [PALETTE[min(i, len(PALETTE)-1)] for i in range(n)]

        bars = ax.barh(range(len(labels)), values,
                       color=bar_colors, edgecolor=HEX_BLANC,
                       height=0.62, linewidth=0.5)

        for bar, val in zip(bars, values):
            ax.text(
                bar.get_width() + max_v * 0.012,
                bar.get_y() + bar.get_height() / 2,
                fmt_aum(val),
                va="center", ha="left",
                fontsize=7.5, color=HEX_MARINE, fontweight="bold"
            )

        ax.set_yticks(range(len(labels)))
        ax.set_yticklabels(
            [l[:24] for l in labels],
            fontsize=8, color=HEX_MARINE
        )
        ax.invert_yaxis()
        ax.xaxis.set_major_formatter(
            plt.FuncFormatter(lambda x, _: fmt_aum(x))
        )
        ax.tick_params(axis="x", labelsize=7, colors=HEX_MARINE)
        ax.set_xlim(0, max_v * 1.28)
        ax.grid(axis="x", alpha=0.22, color=HEX_GRIS, linewidth=0.4)
        ax.spines["left"].set_color(HEX_GRIS)
        ax.spines["bottom"].set_color(HEX_GRIS)
        ax.set_title("Top 10 Deals — AUM Finance",
                     fontsize=9.5, fontweight="bold",
                     color=HEX_MARINE, pad=9)

    fig.tight_layout(pad=0.8)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150,
                bbox_inches="tight", facecolor=HEX_BLANC)
    plt.close(fig)
    buf.seek(0)
    return buf


def _chart_nav_base100(nav_base100_df) -> io.BytesIO:
    """
    Courbe NAV Base 100 — style factsheet minimaliste.
    Design epure : pas de grille verticale, spines reduits, palette Amundi.
    Matplotlib Agg uniquement — aucune dependance Plotly.
    """
    plt.rcParams.update(MPL_PARAMS)
    fig_w = CHART_W / 72
    fig_h = fig_w * 0.46
    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    fig.patch.set_facecolor(HEX_BLANC)
    ax.set_facecolor(HEX_BLANC)

    if nav_base100_df is None or (hasattr(nav_base100_df, "empty") and nav_base100_df.empty):
        ax.text(0.5, 0.5, "Donnees NAV non disponibles pour ce perimetre",
                ha="center", va="center",
                fontsize=9, color=HEX_MARINE, transform=ax.transAxes)
        ax.axis("off")
    else:
        plotted = 0
        for i, fonds in enumerate(nav_base100_df.columns):
            series = nav_base100_df[fonds].dropna()
            if series.empty:
                continue
            color    = PALETTE[i % len(PALETTE)]
            line_sty = "-" if i % 2 == 0 else "--"

            if len(series) >= 2:
                ax.plot(
                    series.index, series.values,
                    label=fonds, color=color,
                    linewidth=1.6, linestyle=line_sty, alpha=0.92
                )
            else:
                ax.scatter(
                    series.index, series.values,
                    color=color, s=55, label=fonds, zorder=5
                )
            # Point terminal
            ax.scatter(
                [series.index[-1]], [series.values[-1]],
                color=color, s=28, zorder=6
            )
            plotted += 1

        ax.axhline(100, color=HEX_GRIS, linewidth=0.7, linestyle=":")
        ax.set_ylabel("Base 100", fontsize=8, color=HEX_MARINE)
        ax.tick_params(colors=HEX_MARINE, labelsize=7.5)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_color(HEX_GRIS)
        ax.spines["bottom"].set_color(HEX_GRIS)
        ax.grid(axis="y", alpha=0.20, color=HEX_GRIS, linewidth=0.4)
        # Pas de grille verticale — minimalisme factsheet
        ax.grid(axis="x", visible=False)

        if plotted > 0:
            ax.legend(
                fontsize=7.5, frameon=True, framealpha=0.92,
                edgecolor=HEX_GRIS, labelcolor=HEX_MARINE,
                loc="upper left",
                ncol=min(3, plotted)
            )
        ax.set_title("Evolution NAV — Base 100",
                     fontsize=9.5, fontweight="bold",
                     color=HEX_MARINE, pad=9)
        plt.xticks(rotation=15, ha="right", fontsize=7)

    fig.tight_layout(pad=0.7)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150,
                bbox_inches="tight", facecolor=HEX_BLANC)
    plt.close(fig)
    buf.seek(0)
    return buf


# ---------------------------------------------------------------------------
# STYLES REPORTLAB
# ---------------------------------------------------------------------------

def _build_styles() -> dict:
    def ps(name, **kw):
        return ParagraphStyle(name, **kw)

    return {
        # Page de garde
        "cover_inner":   ps("cover_inner",
            fontName="Helvetica", fontSize=9,
            textColor=BLANC, leading=17),

        # Titres de sections — Orange Amundi (standard factsheet)
        "section":       ps("section",
            fontName="Helvetica-Bold", fontSize=11.5,
            textColor=ORANGE_AMUNDI,
            spaceBefore=12, spaceAfter=5),

        "section_blue":  ps("section_blue",
            fontName="Helvetica-Bold", fontSize=11.5,
            textColor=BLEU_MARINE,
            spaceBefore=10, spaceAfter=4),

        # Corps de texte
        "body":          ps("body",
            fontName="Helvetica", fontSize=8.5,
            textColor=GRIS_TEXTE,
            spaceAfter=4, leading=13),

        # En-tetes de tableaux
        "th":            ps("th",
            fontName="Helvetica-Bold", fontSize=7.5,
            textColor=BLANC, alignment=TA_LEFT),
        "th_r":          ps("th_r",
            fontName="Helvetica-Bold", fontSize=7.5,
            textColor=BLANC, alignment=TA_RIGHT),

        # Cellules de tableaux
        "td":            ps("td",
            fontName="Helvetica", fontSize=7.5,
            textColor=BLEU_MARINE, alignment=TA_LEFT),
        "td_r":          ps("td_r",
            fontName="Helvetica", fontSize=7.5,
            textColor=BLEU_MARINE, alignment=TA_RIGHT),
        "td_grey":       ps("td_grey",
            fontName="Helvetica", fontSize=7.5,
            textColor=GRIS_TEXTE, alignment=TA_LEFT),
        "td_grey_r":     ps("td_grey_r",
            fontName="Helvetica", fontSize=7.5,
            textColor=GRIS_TEXTE, alignment=TA_RIGHT),

        # KPI page de garde — texte BLANC sur fond marine
        "kpi_label":     ps("kpi_label",
            fontName="Helvetica", fontSize=7,
            textColor=BLANC, alignment=TA_CENTER),
        "kpi_value":     ps("kpi_value",
            fontName="Helvetica-Bold", fontSize=15,
            textColor=BLANC, alignment=TA_CENTER),

        # Disclaimer
        "disclaimer":    ps("disclaimer",
            fontName="Helvetica-Oblique", fontSize=6.5,
            textColor=GRIS_TEXTE, alignment=TA_CENTER, leading=10),

        # Tableau de performances
        "perf_th":       ps("perf_th",
            fontName="Helvetica-Bold", fontSize=7.5,
            textColor=BLANC, alignment=TA_RIGHT),
        "perf_th_l":     ps("perf_th_l",
            fontName="Helvetica-Bold", fontSize=7.5,
            textColor=BLANC, alignment=TA_LEFT),
        "perf_td":       ps("perf_td",
            fontName="Helvetica", fontSize=7.5,
            textColor=BLEU_MARINE, alignment=TA_RIGHT),
        "perf_td_l":     ps("perf_td_l",
            fontName="Helvetica", fontSize=7.5,
            textColor=BLEU_MARINE, alignment=TA_LEFT),
        "perf_pos":      ps("perf_pos",
            fontName="Helvetica-Bold", fontSize=7.5,
            textColor=HexColor("#1A7A3C"), alignment=TA_RIGHT),
        "perf_neg":      ps("perf_neg",
            fontName="Helvetica-Bold", fontSize=7.5,
            textColor=HexColor("#8B2020"), alignment=TA_RIGHT),

        # Legendes graphiques
        "chart_caption": ps("chart_caption",
            fontName="Helvetica-Oblique", fontSize=7,
            textColor=GRIS_TEXTE, alignment=TA_CENTER, spaceAfter=4),

        # Alerte retard
        "alert_retard":  ps("alert_retard",
            fontName="Helvetica-Bold", fontSize=7.5,
            textColor=ORANGE_AMUNDI, alignment=TA_LEFT),
    }


# ---------------------------------------------------------------------------
# PAGE DE GARDE — Fond Marine, Texte Blanc
# ---------------------------------------------------------------------------

def _page_garde(
    styles: dict,
    mode_comex: bool,
    kpis: dict,
    fonds_perimetre: Optional[list] = None
) -> list:
    elements = []
    today_str     = date.today().strftime("%d %B %Y")
    perimetre_str = (
        ", ".join(fonds_perimetre) if fonds_perimetre else "Tous les fonds"
    )

    # Bloc titre — fond marine integre
    cover_data = [[Paragraph(
        f"<font color='#FF4F00' size='7'>"
        f"CONFIDENTIEL"
        f"{'  |  MODE COMEX ACTIF' if mode_comex else ''}"
        f"</font><br/><br/>"
        f"<font color='white' size='22'><b>Executive Report</b></font><br/>"
        f"<font color='white' size='15'>Pipeline &amp; Reporting</font><br/><br/>"
        f"<font color='#B0C8DC' size='8.5'>Asset Management Division</font><br/>"
        f"<font color='white' size='7.5'>Perimetre : {perimetre_str}</font><br/><br/>"
        f"<font color='#B0C8DC' size='7.5'>{today_str}</font>",
        styles["cover_inner"]
    )]]
    cover_tbl = Table(cover_data, colWidths=[USABLE_W])
    cover_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), BLEU_MARINE),
        ("TOPPADDING",    (0, 0), (-1, -1), 44),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 52),
        ("LEFTPADDING",   (0, 0), (-1, -1), 30),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 20),
    ]))
    elements.append(cover_tbl)
    elements.append(Spacer(1, 2))
    # Filet Orange Amundi sous le titre
    elements.append(ColorRect(USABLE_W, 3.5, ORANGE_AMUNDI))
    elements.append(Spacer(1, 16))

    # KPI cards — fond marine, texte blanc
    kpi_items = [
        ("AUM Finance Total", fmt_aum(kpis.get("total_funded", 0))),
        ("Pipeline Actif",    fmt_aum(kpis.get("pipeline_actif", 0))),
        ("Taux Conversion",   f"{kpis.get('taux_conversion', 0):.1f}%"),
        ("Deals Actifs",      str(kpis.get("nb_deals_actifs", 0))),
    ]
    col_w = USABLE_W / 4

    kpi_tbl = Table(
        [[Paragraph(i[0], styles["kpi_label"])  for i in kpi_items],
         [Paragraph(i[1], styles["kpi_value"])  for i in kpi_items]],
        colWidths=[col_w] * 4,
        rowHeights=[0.9 * cm, 1.15 * cm]
    )
    kpi_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), BLEU_MARINE),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",    (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ("LEFTPADDING",   (0, 0), (-1, -1), 4),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 4),
        # Separateurs verticaux fins en orange
        ("LINEAFTER",     (0, 0), (2, -1),  0.4, ORANGE_AMUNDI),
    ]))
    elements.append(kpi_tbl)
    elements.append(Spacer(1, 10))

    disclaimer = (
        "Document strictement confidentiel a usage interne exclusif. "
        "Reproduction et diffusion externe interdites. "
        "Les performances passees ne prejudgent pas des performances futures."
    )
    if mode_comex:
        disclaimer += " Mode Comex actif : noms clients anonymises."
    elements.append(Paragraph(disclaimer, styles["disclaimer"]))
    elements.append(PageBreak())
    return elements


# ---------------------------------------------------------------------------
# SECTION GRAPHIQUES ANALYTIQUES
# Ordre impose : 1.Type Client | 2.Region | 3.Top 10 Deals
# Empiles verticalement, CHART_W = USABLE_W * 0.65, centres
# ---------------------------------------------------------------------------

def _section_charts(
    kpis: dict,
    aum_by_region: dict,
    top_deals_safe: list,
    styles: dict,
    mode_comex: bool,
) -> list:
    elements = []

    # Titre de section — Orange Amundi
    elements.append(Paragraph("Analyse du Pipeline", styles["section"]))
    elements.append(ColorRect(USABLE_W, 2.5, ORANGE_AMUNDI))
    elements.append(Spacer(1, 14))

    # 1. Type de Client
    buf_type = _chart_pie_aum_by_type(kpis.get("aum_by_type", {}))
    _centrer_chart(elements, buf_type, 0.78,
                   "Figure 1 — Repartition AUM Finance par Type de Client",
                   styles)

    # 2. Region Geographique
    buf_reg = _chart_pie_aum_by_region(aum_by_region)
    _centrer_chart(elements, buf_reg, 0.78,
                   "Figure 2 — Repartition AUM Finance par Region Geographique",
                   styles)

    # 3. Top 10 Deals — aspect dynamique selon nombre de deals
    n_deals = len(top_deals_safe[:10])
    aspect  = max(0.42, n_deals * 0.058 + 0.18)
    buf_top = _chart_top10_deals(top_deals_safe, mode_comex)
    _centrer_chart(elements, buf_top, aspect,
                   "Figure 3 — Top 10 Deals par AUM Finance",
                   styles)

    return elements


# ---------------------------------------------------------------------------
# SECTION TABLEAU PIPELINE ACTIF
# ---------------------------------------------------------------------------

def _section_pipeline(
    pipeline_df: pd.DataFrame,
    styles: dict,
    mode_comex: bool
) -> list:
    elements = []
    elements.append(PageBreak())
    elements.append(Paragraph("Pipeline Actif — Recapitulatif", styles["section"]))
    elements.append(ColorRect(USABLE_W, 2.5, ORANGE_AMUNDI))
    elements.append(Spacer(1, 10))

    active = pipeline_df[pipeline_df["statut"].isin(
        ["Prospect", "Initial Pitch", "Due Diligence", "Soft Commit"]
    )].copy()

    if active.empty:
        elements.append(Paragraph(
            "Aucun deal actif dans le perimetre.",
            styles["body"]
        ))
    else:
        if mode_comex:
            active["nom_client"] = active.apply(
                lambda r: f"{r['type_client']} — {r['region']}", axis=1
            )

        # Ratios colonnes — somme = 1.0 exactement
        ratios     = [0.22, 0.16, 0.12, 0.13, 0.13, 0.13, 0.11]
        col_widths = [USABLE_W * r for r in ratios]
        headers    = ["Client", "Fonds", "Statut",
                      "AUM Cible", "AUM Revise",
                      "Prochaine Action", "Commercial"]

        rows  = [[Paragraph(h, styles["th"]) for h in headers]]
        today = date.today()

        for _, row in active.iterrows():
            nad = row.get("next_action_date")
            if isinstance(nad, date):
                overdue  = nad < today
                nad_str  = f"[!] {nad.isoformat()}" if overdue else nad.isoformat()
                nad_style = styles["alert_retard"] if overdue else styles["td"]
            else:
                nad_str   = "—"
                nad_style = styles["td"]

            rows.append([
                Paragraph(str(row.get("nom_client", ""))[:26],       styles["td"]),
                Paragraph(str(row.get("fonds", "")),                  styles["td"]),
                Paragraph(str(row.get("statut", "")),                 styles["td"]),
                Paragraph(fmt_aum(float(row.get("target_aum_initial", 0) or 0)),
                          styles["td_r"]),
                Paragraph(fmt_aum(float(row.get("revised_aum", 0) or 0)),
                          styles["td_r"]),
                Paragraph(nad_str,                                    nad_style),
                Paragraph(str(row.get("sales_owner", ""))[:14],      styles["td"]),
            ])

        tbl = Table(rows, colWidths=col_widths, repeatRows=1)
        tbl.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (-1, 0),  BLEU_MARINE),
            ("LINEBELOW",     (0, 0), (-1, 0),  1.5, ORANGE_AMUNDI),
            ("GRID",          (0, 0), (-1, -1), 0.3, GRIS_CLAIR),
            ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING",    (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING",   (0, 0), (-1, -1), 5),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 5),
            *[("BACKGROUND",  (0, i), (-1, i),  HexColor("#F7FAFD"))
              for i in range(2, len(rows), 2)],
        ]))
        elements.append(tbl)

    # Tableau Lost / Paused
    lost_paused = pipeline_df[
        pipeline_df["statut"].isin(["Lost", "Paused"])
    ].copy()

    if not lost_paused.empty:
        elements.append(Spacer(1, 16))
        elements.append(Paragraph("Deals Perdus / En Pause", styles["section"]))
        elements.append(ColorRect(USABLE_W, 1.5, GRIS_CLAIR))
        elements.append(Spacer(1, 8))

        if mode_comex:
            lost_paused["nom_client"] = lost_paused.apply(
                lambda r: f"{r['type_client']} — {r['region']}", axis=1
            )

        lp_ratios  = [0.26, 0.18, 0.12, 0.22, 0.22]
        lp_col_w   = [USABLE_W * r for r in lp_ratios]
        lp_headers = ["Client", "Fonds", "Statut", "Raison", "Concurrent"]
        lp_rows    = [[Paragraph(h, styles["th"]) for h in lp_headers]]

        for _, row in lost_paused.iterrows():
            lp_rows.append([
                Paragraph(str(row.get("nom_client", ""))[:24],            styles["td_grey"]),
                Paragraph(str(row.get("fonds", "")),                      styles["td_grey"]),
                Paragraph(str(row.get("statut", "")),                     styles["td_grey"]),
                Paragraph(str(row.get("raison_perte", "") or "—"),        styles["td_grey"]),
                Paragraph(str(row.get("concurrent_choisi", "") or "—"),   styles["td_grey"]),
            ])

        lp_tbl = Table(lp_rows, colWidths=lp_col_w, repeatRows=1)
        lp_tbl.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (-1, 0),  HexColor("#7A7A7A")),
            ("LINEBELOW",     (0, 0), (-1, 0),  1.0, GRIS_CLAIR),
            ("GRID",          (0, 0), (-1, -1), 0.3, GRIS_CLAIR),
            ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING",    (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING",   (0, 0), (-1, -1), 5),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 5),
        ]))
        elements.append(lp_tbl)

    return elements


# ---------------------------------------------------------------------------
# SECTION PERFORMANCE (optionnelle)
# Page Performance avec graphique NAV Base 100 + tableau de perf filtre
# ---------------------------------------------------------------------------

def _section_performance(
    perf_data: pd.DataFrame,
    nav_base100_df,
    styles: dict,
    fonds_perimetre: Optional[list] = None,
) -> list:
    """
    Page Performance. Filtre perf_data et nav_base100_df sur fonds_perimetre.
    Design factsheet : minimaliste, palette Amundi, titres Orange.
    """
    elements = []
    elements.append(PageBreak())
    elements.append(Paragraph("Analyse de Performance — NAV", styles["section"]))
    elements.append(ColorRect(USABLE_W, 2.5, ORANGE_AMUNDI))
    elements.append(Spacer(1, 12))

    # Filtrage sur le perimetre
    pf = perf_data.copy() if perf_data is not None else pd.DataFrame()
    nb = nav_base100_df.copy() if nav_base100_df is not None else None

    if fonds_perimetre:
        if not pf.empty and "Fonds" in pf.columns:
            pf = pf[pf["Fonds"].isin(fonds_perimetre)]
        if nb is not None and not (hasattr(nb, "empty") and nb.empty):
            try:
                cols_keep = [c for c in nb.columns if c in fonds_perimetre]
                nb = nb[cols_keep] if cols_keep else pd.DataFrame()
            except Exception:
                nb = pd.DataFrame()

    # Graphique NAV Base 100
    buf_nav = _chart_nav_base100(nb)
    _centrer_chart(elements, buf_nav, 0.50,
                   "Figure 4 — Evolution NAV Base 100 — Periode selectionnee",
                   styles)

    if pf.empty:
        elements.append(Paragraph(
            "Aucune donnee de performance disponible pour ce perimetre.",
            styles["body"]
        ))
        return elements

    # Tableau des performances
    elements.append(Paragraph("Tableau des Performances", styles["section"]))
    elements.append(ColorRect(USABLE_W, 2.5, ORANGE_AMUNDI))
    elements.append(Spacer(1, 10))

    perf_col_names = {"Perf 1M (%)", "Perf YTD (%)", "Perf Periode (%)"}
    col_order  = ["Fonds", "NAV Derniere", "Base 100 Actuel",
                  "Perf 1M (%)", "Perf YTD (%)", "Perf Periode (%)"]
    available  = [c for c in col_order if c in pf.columns]
    n          = len(available)

    if n == 0:
        elements.append(Paragraph(
            "Colonnes manquantes dans les donnees de performance.",
            styles["body"]
        ))
        return elements

    fonds_r    = 0.28
    rest_r     = (1.0 - fonds_r) / max(n - 1, 1)
    col_widths = [USABLE_W * (fonds_r if i == 0 else rest_r)
                  for i in range(n)]

    h_row  = [
        Paragraph(c, styles["perf_th_l"] if i == 0 else styles["perf_th"])
        for i, c in enumerate(available)
    ]
    rows = [h_row]

    def _fmt_p(val):
        if val is None or (isinstance(val, float) and np.isnan(val)):
            return "n.d."
        sign = "+" if val > 0 else ""
        return f"{sign}{val:.2f}%"

    def _p_style(val):
        if val is None or (isinstance(val, float) and np.isnan(val)):
            return styles["perf_td"]
        return styles["perf_pos"] if val >= 0 else styles["perf_neg"]

    for _, row in pf.iterrows():
        data_row = []
        for i, col in enumerate(available):
            val = row.get(col)
            if col == "Fonds":
                data_row.append(Paragraph(str(val), styles["perf_td_l"]))
            elif col in perf_col_names:
                fval = float(val) if (val is not None and val == val) else None
                data_row.append(Paragraph(_fmt_p(fval), _p_style(fval)))
            else:
                try:
                    fval = float(val)
                    txt  = f"{fval:.4f}" if "NAV" in col else f"{fval:.2f}"
                except Exception:
                    txt = "—"
                data_row.append(Paragraph(txt, styles["perf_td"]))
        rows.append(data_row)

    perf_tbl = Table(rows, colWidths=col_widths, repeatRows=1)
    perf_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, 0),  BLEU_MARINE),
        ("LINEBELOW",     (0, 0), (-1, 0),  1.5, ORANGE_AMUNDI),
        ("GRID",          (0, 0), (-1, -1), 0.3, GRIS_CLAIR),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",    (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING",   (0, 0), (-1, -1), 6),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 6),
        *[("BACKGROUND",  (0, i), (-1, i),  HexColor("#F7FAFD"))
          for i in range(2, len(rows), 2)],
    ]))
    elements.append(perf_tbl)
    return elements


# ---------------------------------------------------------------------------
# EN-TETE / PIED DE PAGE
# ---------------------------------------------------------------------------

def _header_footer(canvas, doc):
    canvas.saveState()
    if doc.page > 1:
        # Bande de pied marine
        canvas.setFillColor(BLEU_MARINE)
        canvas.rect(0, 0, PAGE_W, 0.95 * cm, fill=1, stroke=0)
        canvas.setFont("Helvetica", 6.5)
        canvas.setFillColor(HexColor("#B0C8DC"))
        canvas.drawString(
            2 * cm, 0.32 * cm,
            "Executive Report — Asset Management Division"
        )
        canvas.drawRightString(
            PAGE_W - 2 * cm, 0.32 * cm,
            f"Page {doc.page}"
        )
        # Filet orange en haut de page
        canvas.setStrokeColor(ORANGE_AMUNDI)
        canvas.setLineWidth(2.0)
        canvas.line(0, PAGE_H - 0.40 * cm, PAGE_W, PAGE_H - 0.40 * cm)
        canvas.setFont("Helvetica", 6.5)
        canvas.setFillColor(BLEU_MARINE)
        canvas.drawRightString(
            PAGE_W - 2 * cm, PAGE_H - 0.28 * cm,
            "CONFIDENTIEL"
        )
    canvas.restoreState()


# ---------------------------------------------------------------------------
# FONCTION PRINCIPALE generate_pdf
# ---------------------------------------------------------------------------

def generate_pdf(
    pipeline_df: pd.DataFrame,
    kpis: dict,
    aum_by_region: Optional[dict] = None,
    mode_comex: bool = False,
    perf_data=None,
    nav_base100_df=None,
    fonds_perimetre: Optional[list] = None,
) -> bytes:
    """
    Genere le rapport PDF Executive Pitchbook complet.

    Args:
        pipeline_df    : DataFrame pipeline filtre sur fonds_perimetre.
        kpis           : Dict KPIs deja calcule sur fonds_perimetre.
        aum_by_region  : Dict AUM par region (deja filtre).
        mode_comex     : Anonymisation totale des noms clients.
        perf_data      : DataFrame performances NAV (optionnel).
        nav_base100_df : DataFrame pivot Base100 (optionnel).
        fonds_perimetre: Liste des fonds du perimetre (pour en-tetes PDF).

    Returns:
        bytes du PDF pret au telechargement.
    """
    # --- Anonymisation Mode Comex ---
    if mode_comex:
        pipeline_df = anonymize_df(pipeline_df)
        top_deals_safe = [
            {**d, "nom_client": f"{d['type_client']} — {d['region']}"}
            for d in kpis.get("top_deals", [])
        ]
        kpis = {**kpis, "top_deals": top_deals_safe}
    else:
        top_deals_safe = kpis.get("top_deals", [])

    aum_by_region = aum_by_region or {}
    styles        = _build_styles()
    pdf_buf       = io.BytesIO()

    doc = SimpleDocTemplate(
        pdf_buf,
        pagesize=A4,
        leftMargin=MARGIN_H,  rightMargin=MARGIN_H,
        topMargin=MARGIN_V,   bottomMargin=1.5 * cm,
        title="Executive Report — Asset Management",
        author="Asset Management Division",
    )

    elements = []

    # Page 1 : Couverture + KPIs
    elements += _page_garde(styles, mode_comex, kpis, fonds_perimetre)

    # Page 2 : Graphiques empiles (Type Client, Region, Top 10)
    elements += _section_charts(
        kpis, aum_by_region, top_deals_safe, styles, mode_comex
    )

    # Page 3 : Tableau Pipeline
    elements += _section_pipeline(pipeline_df, styles, mode_comex)

    # Page 4 (optionnelle) : Performance & NAV
    if perf_data is not None and not (hasattr(perf_data, "empty") and perf_data.empty):
        elements += _section_performance(
            perf_data, nav_base100_df, styles, fonds_perimetre
        )

    doc.build(
        elements,
        onFirstPage=_header_footer,
        onLaterPages=_header_footer
    )

    result = pdf_buf.getvalue()
    pdf_buf.close()
    return result
