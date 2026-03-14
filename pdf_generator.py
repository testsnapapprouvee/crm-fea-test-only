# =============================================================================
# pdf_generator.py  —  CRM Asset Management  —  Amundi Edition
# Charte : #001c4b Marine | #019ee1 Ciel | #f07d00 Orange
# Design Research :
#   - Titres de section et sous-section en Orange #f07d00
#   - Graphiques à 65% de la largeur utile, centrés
#   - KPI page de garde : padding généreux, effet aéré
#   - Top 10 : page dédiée, Inflows + Outflows optionnels
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
    Image, PageBreak, KeepTogether,
)
from reportlab.platypus.flowables import Flowable


# ---------------------------------------------------------------------------
# COULEURS
# ---------------------------------------------------------------------------
COL_MARINE = HexColor("#001c4b")
COL_CIEL   = HexColor("#019ee1")
COL_ORANGE = HexColor("#f07d00")
COL_BLANC  = HexColor("#ffffff")
COL_GRIS   = HexColor("#e8e8e8")
COL_TEXTE  = HexColor("#444444")
COL_VERT   = HexColor("#1a7a3c")
COL_ROUGE  = HexColor("#8b2020")
COL_HEADER = HexColor("#f0f5fa")

HX_MARINE = "#001c4b"
HX_CIEL   = "#019ee1"
HX_ORANGE = "#f07d00"
HX_BLANC  = "#ffffff"
HX_GRIS   = "#e8e8e8"
HX_TEXTE  = "#444444"

# ---------------------------------------------------------------------------
# MISE EN PAGE
# ---------------------------------------------------------------------------
PAGE_W, PAGE_H = A4
MARGIN_H  = 2.0 * cm
MARGIN_V  = 1.8 * cm
USABLE_W  = PAGE_W - 2 * MARGIN_H          # 481.9 pts

# Proportions Design Research : graphiques à 65%, centrés
CHART_W   = USABLE_W * 0.65                # 313.2 pts — largeur graphique
CHART_PAD = (USABLE_W - CHART_W) / 2.0     # marge de centrage

# Donuts : 65% répartis en 2 colonnes égales
DONUT_W   = USABLE_W * 0.325               # chaque donut = la moitié de 65%
DONUT_H   = DONUT_W  * 1.05               # légèrement plus haut que large

# Top 10 : 65% centré
TOP10_W   = CHART_W
TOP10_PAD = CHART_PAD

# NAV : légèrement plus large (80%) pour la lisibilité des courbes
NAV_W     = USABLE_W * 0.80
NAV_H     = NAV_W * 0.42
NAV_PAD   = (USABLE_W - NAV_W) / 2.0

PALETTE = [
    "#1a5e8a", "#001c4b", "#4a8fbd", "#003f7a",
    "#2c7fb8", "#004f8c", "#6baed6", "#08519c",
    "#9ecae1", "#003060",
]

MPL_RC = {
    "font.family":       "DejaVu Sans",
    "font.size":         8.5,
    "axes.spines.top":   False,
    "axes.spines.right": False,
    "axes.edgecolor":    HX_GRIS,
    "axes.linewidth":    0.6,
    "axes.labelcolor":   HX_MARINE,
    "xtick.color":       HX_MARINE,
    "ytick.color":       HX_MARINE,
    "figure.facecolor":  HX_BLANC,
    "axes.facecolor":    HX_BLANC,
    "grid.color":        HX_GRIS,
    "grid.linewidth":    0.4,
}


# ---------------------------------------------------------------------------
# FLOWABLE
# ---------------------------------------------------------------------------
class ColorRect(Flowable):
    def __init__(self, width, height, color):
        super().__init__()
        self.width = width
        self.height = height
        self.color = color

    def draw(self):
        self.canv.setFillColor(self.color)
        self.canv.rect(0, 0, self.width, self.height, fill=1, stroke=0)


# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------
def fmt_aum(value):
    try:
        v = float(value)
    except (TypeError, ValueError):
        return "-"
    if v == 0:
        return "0.0 M EUR"
    if v >= 1_000_000_000:
        return "{:.1f} Md EUR".format(v / 1_000_000_000)
    return "{:.1f} M EUR".format(v / 1_000_000)


def _center_image(img, img_w, img_h, pad):
    """Encapsule une image dans un tableau centré (padding gauche/droite)."""
    tbl = Table(
        [[Spacer(pad, 1), img, Spacer(pad, 1)]],
        colWidths=[pad, img_w, pad],
    )
    tbl.setStyle(TableStyle([
        ("LEFTPADDING",   (0, 0), (-1, -1), 0),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 0),
        ("TOPPADDING",    (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    return tbl


# ---------------------------------------------------------------------------
# GRAPHIQUES MATPLOTLIB
# ---------------------------------------------------------------------------

def _make_donut_png(labels, values, title, fig_w_in, fig_h_in):
    plt.rcParams.update(MPL_RC)
    fig, ax = plt.subplots(figsize=(fig_w_in, fig_h_in))
    fig.patch.set_facecolor(HX_BLANC)
    ax.set_facecolor(HX_BLANC)

    if not labels or not values or sum(values) == 0:
        ax.text(0.5, 0.5, "Aucune donnee", ha="center", va="center",
                fontsize=8, color=HX_MARINE, transform=ax.transAxes)
        ax.axis("off")
    else:
        colors = [PALETTE[i % len(PALETTE)] for i in range(len(labels))]
        _, _, autotexts = ax.pie(
            values, colors=colors, autopct="%1.0f%%", startangle=90,
            pctdistance=0.72,
            wedgeprops={"width": 0.52, "edgecolor": HX_BLANC, "linewidth": 1.5},
        )
        for at in autotexts:
            at.set_fontsize(7.5)
            at.set_color(HX_BLANC)
            at.set_fontweight("bold")

        total = sum(values)
        ax.text(0,  0.10, fmt_aum(total), ha="center", va="center",
                fontsize=8, fontweight="bold", color=HX_MARINE)
        ax.text(0, -0.15, "Finance", ha="center", va="center",
                fontsize=6.5, color=HX_TEXTE)

        patches = [
            mpatches.Patch(color=colors[i],
                           label="{}: {}".format(labels[i], fmt_aum(values[i])))
            for i in range(len(labels))
        ]
        ax.legend(handles=patches, loc="lower center",
                  bbox_to_anchor=(0.5, -0.32), ncol=2,
                  fontsize=6.5, frameon=False, labelcolor=HX_MARINE)
        ax.set_title(title, fontsize=8.5, fontweight="bold",
                     color=HX_ORANGE, pad=7)   # titre donut en orange

    fig.subplots_adjust(left=0.02, right=0.98, top=0.88, bottom=0.26)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, facecolor=HX_BLANC)
    plt.close(fig)
    buf.seek(0)
    return buf


def _make_top10_png(deals, mode_comex, fig_w_in, fig_h_in,
                    title="Top 10 Inflows — AUM Finance"):
    plt.rcParams.update(MPL_RC)
    fig, ax = plt.subplots(figsize=(fig_w_in, fig_h_in))
    fig.patch.set_facecolor(HX_BLANC)
    ax.set_facecolor(HX_BLANC)

    deals_10 = deals[:10]
    if not deals_10:
        ax.text(0.5, 0.5, "Aucun deal dans cette categorie",
                ha="center", va="center", fontsize=9, color=HX_MARINE,
                transform=ax.transAxes)
        ax.axis("off")
    else:
        labels = [
            "{} - {}".format(d.get("type_client", ""), d.get("region", ""))
            if mode_comex else str(d.get("nom_client", ""))
            for d in deals_10
        ]
        values = [float(d.get("funded_aum", 0)) for d in deals_10]
        pairs  = sorted(zip(values, labels), reverse=True)
        values = [v for v, _ in pairs]
        labels = [l for _, l in pairs]
        max_v  = max(values) if values else 1.0

        colors = [HX_ORANGE if i == 0 else HX_MARINE for i in range(len(deals_10))]
        bars   = ax.barh(range(len(labels)), values,
                         color=colors, edgecolor=HX_BLANC,
                         height=0.65, linewidth=0.3)

        for bar, val in zip(bars, values):
            ax.text(
                bar.get_width() + max_v * 0.015,
                bar.get_y() + bar.get_height() / 2,
                fmt_aum(val),
                va="center", ha="left",
                fontsize=8, color=HX_MARINE, fontweight="bold",
            )

        ax.set_yticks(range(len(labels)))
        ax.set_yticklabels([l[:28] for l in labels], fontsize=8, color=HX_MARINE)
        ax.invert_yaxis()
        ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: fmt_aum(x)))
        ax.tick_params(axis="x", labelsize=7.5, colors=HX_MARINE)
        ax.set_xlim(0, max_v * 1.32)
        ax.grid(axis="x", alpha=0.22, color=HX_GRIS, linewidth=0.35)
        ax.spines["left"].set_color(HX_GRIS)
        ax.spines["bottom"].set_color(HX_GRIS)
        ax.set_title(title, fontsize=10, fontweight="bold",
                     color=HX_ORANGE, pad=10)  # titre top10 en orange

    fig.subplots_adjust(left=0.26, right=0.86, top=0.92, bottom=0.06)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, facecolor=HX_BLANC)
    plt.close(fig)
    buf.seek(0)
    return buf


def _make_nav_png(nav_df, fig_w_in, fig_h_in):
    plt.rcParams.update(MPL_RC)
    fig, ax = plt.subplots(figsize=(fig_w_in, fig_h_in))
    fig.patch.set_facecolor(HX_BLANC)
    ax.set_facecolor(HX_BLANC)

    if nav_df is None or not hasattr(nav_df, "columns") or nav_df.empty:
        ax.text(0.5, 0.5, "Donnees NAV non disponibles",
                ha="center", va="center", fontsize=9, color=HX_MARINE,
                transform=ax.transAxes)
        ax.axis("off")
    else:
        NAV_COLORS = [
            "#001c4b", "#019ee1", "#1a5e8a", "#4a8fbd",
            "#003f7a", "#2c7fb8", "#004f8c", "#6baed6",
        ]
        for i, col in enumerate(nav_df.columns):
            series = nav_df[col].dropna()
            if series.empty:
                continue
            color = NAV_COLORS[i % len(NAV_COLORS)]
            if len(series) >= 2:
                ax.plot(series.index, series.values,
                        label=col, color=color, linewidth=1.5, alpha=0.92)
            else:
                ax.scatter(series.index, series.values,
                           color=color, s=50, label=col, zorder=5)

        ax.axhline(100, color=HX_GRIS, linewidth=0.7, linestyle="dotted")
        ax.set_ylabel("Base 100", fontsize=7.5, color=HX_MARINE)
        ax.tick_params(colors=HX_MARINE, labelsize=7)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_color(HX_GRIS)
        ax.spines["bottom"].set_color(HX_GRIS)
        ax.grid(axis="y", alpha=0.18, color=HX_GRIS, linewidth=0.35)
        ax.grid(axis="x", visible=False)
        ax.legend(fontsize=7, frameon=True, framealpha=0.90,
                  edgecolor=HX_GRIS, labelcolor=HX_MARINE, loc="upper left")
        ax.set_title("Evolution NAV - Base 100",
                     fontsize=9, fontweight="bold", color=HX_ORANGE, pad=8)
        plt.xticks(rotation=12, ha="right", fontsize=7)

    fig.subplots_adjust(left=0.08, right=0.97, top=0.90, bottom=0.14)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, facecolor=HX_BLANC)
    plt.close(fig)
    buf.seek(0)
    return buf


# ---------------------------------------------------------------------------
# STYLES REPORTLAB
# Titres de section et sous-section : Orange #f07d00 (Design Research)
# ---------------------------------------------------------------------------
def _build_styles():
    s_cover   = ParagraphStyle("cover",   fontName="Helvetica",         fontSize=9,    textColor=COL_BLANC,  leading=17)
    # Titres de section → Orange
    s_section = ParagraphStyle("section", fontName="Helvetica-Bold",    fontSize=11,   textColor=COL_ORANGE, spaceBefore=10, spaceAfter=4)
    s_subsect = ParagraphStyle("subsect", fontName="Helvetica-Bold",    fontSize=9,    textColor=COL_ORANGE, spaceBefore=6,  spaceAfter=3)
    s_body    = ParagraphStyle("body",    fontName="Helvetica",         fontSize=8.5,  textColor=COL_TEXTE,  spaceAfter=4,  leading=13)
    s_th      = ParagraphStyle("th",      fontName="Helvetica-Bold",    fontSize=7.5,  textColor=COL_BLANC,  alignment=TA_LEFT)
    s_th_r    = ParagraphStyle("th_r",    fontName="Helvetica-Bold",    fontSize=7.5,  textColor=COL_BLANC,  alignment=TA_RIGHT)
    s_td      = ParagraphStyle("td",      fontName="Helvetica",         fontSize=7.5,  textColor=COL_MARINE, alignment=TA_LEFT)
    s_td_r    = ParagraphStyle("td_r",    fontName="Helvetica",         fontSize=7.5,  textColor=COL_MARINE, alignment=TA_RIGHT)
    s_td_g    = ParagraphStyle("td_g",    fontName="Helvetica",         fontSize=7.5,  textColor=COL_TEXTE,  alignment=TA_LEFT)
    s_kpi_lbl = ParagraphStyle("kpi_lbl", fontName="Helvetica",         fontSize=7.5,  textColor=COL_BLANC,  alignment=TA_CENTER)
    s_kpi_val = ParagraphStyle("kpi_val", fontName="Helvetica-Bold",    fontSize=16,   textColor=COL_BLANC,  alignment=TA_CENTER)
    s_disc    = ParagraphStyle("disc",    fontName="Helvetica-Oblique", fontSize=6.5,  textColor=COL_TEXTE,  alignment=TA_CENTER, leading=10)
    s_pth_l   = ParagraphStyle("pth_l",  fontName="Helvetica-Bold",    fontSize=7.5,  textColor=COL_BLANC,  alignment=TA_LEFT)
    s_pth_r   = ParagraphStyle("pth_r",  fontName="Helvetica-Bold",    fontSize=7.5,  textColor=COL_BLANC,  alignment=TA_RIGHT)
    s_ptd_l   = ParagraphStyle("ptd_l",  fontName="Helvetica",         fontSize=7.5,  textColor=COL_MARINE, alignment=TA_LEFT)
    s_ptd     = ParagraphStyle("ptd",    fontName="Helvetica",         fontSize=7.5,  textColor=COL_MARINE, alignment=TA_RIGHT)
    s_ppos    = ParagraphStyle("ppos",   fontName="Helvetica-Bold",    fontSize=7.5,  textColor=COL_VERT,   alignment=TA_RIGHT)
    s_pneg    = ParagraphStyle("pneg",   fontName="Helvetica-Bold",    fontSize=7.5,  textColor=COL_ROUGE,  alignment=TA_RIGHT)
    s_caption = ParagraphStyle("caption",fontName="Helvetica-Oblique", fontSize=6.5,  textColor=COL_TEXTE,  alignment=TA_CENTER, spaceAfter=3)
    # Badge retard : orange
    s_alert   = ParagraphStyle("alert",  fontName="Helvetica-Bold",    fontSize=7.5,  textColor=COL_ORANGE, alignment=TA_LEFT)
    return {
        "cover": s_cover, "section": s_section, "subsect": s_subsect,
        "body": s_body, "th": s_th, "th_r": s_th_r,
        "td": s_td, "td_r": s_td_r, "td_g": s_td_g,
        "kpi_lbl": s_kpi_lbl, "kpi_val": s_kpi_val, "disc": s_disc,
        "pth_l": s_pth_l, "pth_r": s_pth_r, "ptd_l": s_ptd_l,
        "ptd": s_ptd, "ppos": s_ppos, "pneg": s_pneg,
        "caption": s_caption, "alert": s_alert,
    }


# ---------------------------------------------------------------------------
# PAGE DE GARDE
# KPI bandeau : padding augmenté pour un rendu aéré (plus de "tassé")
# ---------------------------------------------------------------------------
def _page_garde(styles, mode_comex, kpis, fonds_perimetre):
    elements = []
    today_str = date.today().strftime("%d %B %Y")
    perim_str = ", ".join(fonds_perimetre) if fonds_perimetre else "Tous les fonds"

    cover_text = (
        "<font color='#f07d00' size='7'>CONFIDENTIEL"
        + ("  |  MODE COMEX ACTIF" if mode_comex else "")
        + "</font><br/><br/>"
        "<font color='white' size='22'><b>Executive Report</b></font><br/>"
        "<font color='white' size='15'>Pipeline &amp; Reporting</font><br/><br/>"
        "<font color='#7ab8d8' size='8.5'>Asset Management Division</font><br/>"
        "<font color='white' size='7.5'>Perimetre : " + perim_str + "</font><br/><br/>"
        "<font color='#7ab8d8' size='7.5'>" + today_str + "</font>"
    )
    cover_tbl = Table([[Paragraph(cover_text, styles["cover"])]], colWidths=[USABLE_W])
    cover_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), COL_MARINE),
        ("TOPPADDING",    (0, 0), (-1, -1), 48),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 56),
        ("LEFTPADDING",   (0, 0), (-1, -1), 36),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 24),
    ]))
    elements.append(cover_tbl)
    elements.append(Spacer(1, 2))
    # Filet orange — accent unique
    elements.append(ColorRect(USABLE_W, 4, COL_ORANGE))
    elements.append(Spacer(1, 20))

    # KPI cards — padding généreux, effet aéré
    kpi_items = [
        ("AUM Finance Total", fmt_aum(kpis.get("total_funded", 0))),
        ("Pipeline Actif",    fmt_aum(kpis.get("pipeline_actif", 0))),
        ("Taux Conversion",   "{:.1f}%".format(kpis.get("taux_conversion", 0))),
        ("Deals Actifs",      str(kpis.get("nb_deals_actifs", 0))),
    ]
    col_w = USABLE_W / 4.0
    kpi_tbl = Table(
        [[Paragraph(i[0], styles["kpi_lbl"]) for i in kpi_items],
         [Paragraph(i[1], styles["kpi_val"]) for i in kpi_items]],
        colWidths=[col_w] * 4,
        rowHeights=[1.10 * cm, 1.50 * cm],   # hauteurs augmentées — rendu aéré
    )
    kpi_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), COL_MARINE),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",    (0, 0), (-1, -1), 14),   # padding vertical doublé
        ("BOTTOMPADDING", (0, 0), (-1, -1), 14),
        ("LEFTPADDING",   (0, 0), (-1, -1), 10),   # padding horizontal doublé
        ("RIGHTPADDING",  (0, 0), (-1, -1), 10),
        ("LINEAFTER",     (0, 0), (2, -1),  0.5, COL_ORANGE),   # séparateur orange
    ]))
    elements.append(kpi_tbl)
    elements.append(Spacer(1, 14))

    disc = (
        "Document strictement confidentiel a usage interne exclusif. "
        "Reproduction et diffusion externe interdites. "
        "Les performances passees ne prejudgent pas des performances futures."
    )
    if mode_comex:
        disc += " Mode Comex actif : noms clients anonymises."
    elements.append(Paragraph(disc, styles["disc"]))
    elements.append(PageBreak())
    return elements


# ---------------------------------------------------------------------------
# PAGE 2 : DONUTS — 65% centrés, côte à côte
# ---------------------------------------------------------------------------
def _section_donuts(kpis, aum_by_region, styles):
    elements = []
    # Titre de section en orange
    elements.append(Paragraph("Repartition AUM Finance", styles["section"]))
    elements.append(ColorRect(USABLE_W, 2, COL_ORANGE))
    elements.append(Spacer(1, 16))

    d_w_in = DONUT_W / 72.0
    d_h_in = DONUT_H / 72.0

    abt    = kpis.get("aum_by_type", {})
    buf_d1 = _make_donut_png(list(abt.keys()), list(abt.values()),
                             "AUM par Type de Client", d_w_in, d_h_in)
    img_d1 = Image(buf_d1, width=DONUT_W, height=DONUT_H)

    buf_d2 = _make_donut_png(list(aum_by_region.keys()), list(aum_by_region.values()),
                             "AUM par Region Geographique", d_w_in, d_h_in)
    img_d2 = Image(buf_d2, width=DONUT_W, height=DONUT_H)

    # Centrage : padding de chaque côté = CHART_PAD
    inner_gap = CHART_W - 2 * DONUT_W   # espace entre les deux donuts
    tbl = Table(
        [[Spacer(CHART_PAD, 1), img_d1, Spacer(inner_gap, 1), img_d2, Spacer(CHART_PAD, 1)]],
        colWidths=[CHART_PAD, DONUT_W, inner_gap, DONUT_W, CHART_PAD],
    )
    tbl.setStyle(TableStyle([
        ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING",   (0, 0), (-1, -1), 0),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 0),
        ("TOPPADDING",    (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    cap = Paragraph(
        "Figure 1 — Repartition AUM Finance par Type de Client (gauche) et par Region (droite)",
        styles["caption"]
    )
    elements.append(KeepTogether([tbl, cap]))
    return elements


# ---------------------------------------------------------------------------
# PAGE 3 : TOP 10 INFLOWS — page dédiée, 65% centrés
# ---------------------------------------------------------------------------
def _section_top10(top_deals, outflows, styles, mode_comex, include_outflows):
    elements = []
    elements.append(PageBreak())
    elements.append(Paragraph("Top 10 Deals", styles["section"]))
    elements.append(ColorRect(USABLE_W, 2, COL_ORANGE))
    elements.append(Spacer(1, 18))

    # --- Inflows (Funded) ---
    elements.append(Paragraph("Inflows — AUM Finance (Funded)", styles["subsect"]))
    elements.append(Spacer(1, 10))

    n_in      = max(len(top_deals[:10]), 1)
    t10_h_pts = max(180, min(320, n_in * 27 + 40))
    t10_w_in  = TOP10_W / 72.0
    t10_h_in  = t10_h_pts / 72.0

    buf_in = _make_top10_png(top_deals, mode_comex, t10_w_in, t10_h_in,
                             title="Top 10 Inflows — AUM Finance")
    img_in = Image(buf_in, width=TOP10_W, height=t10_h_pts)
    cap_in = Paragraph(
        "Figure 2 — Top 10 Deals par AUM Finance (statut Funded uniquement)",
        styles["caption"]
    )
    elements.append(KeepTogether([_center_image(img_in, TOP10_W, t10_h_pts, TOP10_PAD), cap_in]))

    # --- Outflows (Redeemed) ---
    if include_outflows and outflows:
        elements.append(Spacer(1, 22))
        elements.append(Paragraph("Outflows — AUM Rachete (Redeemed)", styles["subsect"]))
        elements.append(Spacer(1, 10))

        n_out     = max(len(outflows[:10]), 1)
        out_h_pts = max(120, min(260, n_out * 27 + 40))
        out_h_in  = out_h_pts / 72.0

        buf_out = _make_top10_png(outflows, mode_comex, t10_w_in, out_h_in,
                                  title="Top 10 Outflows — AUM Rachete")
        img_out = Image(buf_out, width=TOP10_W, height=out_h_pts)
        cap_out = Paragraph(
            "Figure 3 — Top 10 Rachats par AUM (statut Redeemed)",
            styles["caption"]
        )
        elements.append(KeepTogether([_center_image(img_out, TOP10_W, out_h_pts, TOP10_PAD), cap_out]))

    return elements


# ---------------------------------------------------------------------------
# PAGE 4 : PIPELINE ACTIF TABLEAU
# ---------------------------------------------------------------------------
def _section_pipeline(pipeline_df, styles, mode_comex):
    elements = []
    elements.append(PageBreak())
    elements.append(Paragraph("Pipeline Actif - Recapitulatif", styles["section"]))
    elements.append(ColorRect(USABLE_W, 2, COL_ORANGE))
    elements.append(Spacer(1, 12))

    actifs = pipeline_df[pipeline_df["statut"].isin(
        ["Prospect", "Initial Pitch", "Due Diligence", "Soft Commit"]
    )].copy()

    if actifs.empty:
        elements.append(Paragraph("Aucun deal actif dans le perimetre.", styles["body"]))
    else:
        if mode_comex:
            actifs["nom_client"] = actifs.apply(
                lambda r: "{} - {}".format(r.get("type_client", ""), r.get("region", "")),
                axis=1
            )
        ratios     = [0.22, 0.16, 0.12, 0.13, 0.13, 0.13, 0.11]
        col_widths = [USABLE_W * r for r in ratios]
        headers    = ["Client", "Fonds", "Statut", "AUM Cible",
                      "AUM Revise", "Prochaine Action", "Commercial"]
        rows       = [[Paragraph(h, styles["th"]) for h in headers]]
        today      = date.today()

        for _, row in actifs.iterrows():
            nad = row.get("next_action_date")
            if isinstance(nad, date):
                if nad < today:
                    nad_str   = "[RETARD] {}".format(nad.isoformat())
                    nad_style = styles["alert"]
                else:
                    nad_str   = nad.isoformat()
                    nad_style = styles["td"]
            else:
                nad_str   = "-"
                nad_style = styles["td"]

            rows.append([
                Paragraph(str(row.get("nom_client", ""))[:26], styles["td"]),
                Paragraph(str(row.get("fonds", "")),            styles["td"]),
                Paragraph(str(row.get("statut", "")),           styles["td"]),
                Paragraph(fmt_aum(float(row.get("target_aum_initial", 0) or 0)), styles["td_r"]),
                Paragraph(fmt_aum(float(row.get("revised_aum", 0) or 0)),         styles["td_r"]),
                Paragraph(nad_str, nad_style),
                Paragraph(str(row.get("sales_owner", ""))[:14], styles["td"]),
            ])

        tbl = Table(rows, colWidths=col_widths, repeatRows=1)
        tbl.setStyle(TableStyle([
            ("BACKGROUND",     (0, 0), (-1, 0),  COL_MARINE),
            ("LINEBELOW",      (0, 0), (-1, 0),  1.5, COL_ORANGE),   # séparateur orange
            ("GRID",           (0, 0), (-1, -1), 0.3, COL_GRIS),
            ("VALIGN",         (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING",     (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING",  (0, 0), (-1, -1), 5),
            ("LEFTPADDING",    (0, 0), (-1, -1), 5),
            ("RIGHTPADDING",   (0, 0), (-1, -1), 5),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [COL_HEADER, COL_BLANC]),
        ]))
        elements.append(tbl)

    # Tableau Lost / Paused
    lost_paused = pipeline_df[pipeline_df["statut"].isin(["Lost", "Paused"])].copy()
    if not lost_paused.empty:
        elements.append(Spacer(1, 18))
        elements.append(Paragraph("Deals Perdus / En Pause", styles["subsect"]))
        elements.append(ColorRect(USABLE_W, 1, COL_GRIS))
        elements.append(Spacer(1, 8))
        if mode_comex:
            lost_paused["nom_client"] = lost_paused.apply(
                lambda r: "{} - {}".format(r.get("type_client", ""), r.get("region", "")),
                axis=1
            )
        lp_ratios  = [0.26, 0.18, 0.12, 0.22, 0.22]
        lp_col_w   = [USABLE_W * r for r in lp_ratios]
        lp_rows    = [[Paragraph(h, styles["th"])
                       for h in ["Client", "Fonds", "Statut", "Raison", "Concurrent"]]]
        for _, row in lost_paused.iterrows():
            lp_rows.append([
                Paragraph(str(row.get("nom_client", ""))[:24],          styles["td_g"]),
                Paragraph(str(row.get("fonds", "")),                    styles["td_g"]),
                Paragraph(str(row.get("statut", "")),                   styles["td_g"]),
                Paragraph(str(row.get("raison_perte", "") or "-"),      styles["td_g"]),
                Paragraph(str(row.get("concurrent_choisi", "") or "-"), styles["td_g"]),
            ])
        lp_tbl = Table(lp_rows, colWidths=lp_col_w, repeatRows=1)
        lp_tbl.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (-1, 0),  HexColor("#7a7a7a")),
            ("GRID",          (0, 0), (-1, -1), 0.3, COL_GRIS),
            ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING",    (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING",   (0, 0), (-1, -1), 5),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 5),
        ]))
        elements.append(lp_tbl)
    return elements


# ---------------------------------------------------------------------------
# PAGE 5 : PERFORMANCE NAV (optionnelle) — graphique 80%, tableau complet
# ---------------------------------------------------------------------------
def _section_performance(perf_data, nav_base100_df, styles, fonds_perimetre):
    elements = []
    elements.append(PageBreak())
    elements.append(Paragraph("Analyse de Performance - NAV", styles["section"]))
    elements.append(ColorRect(USABLE_W, 2, COL_ORANGE))
    elements.append(Spacer(1, 14))

    pf = perf_data.copy() if perf_data is not None else pd.DataFrame()
    nb = nav_base100_df.copy() if nav_base100_df is not None else None

    if fonds_perimetre:
        if not pf.empty and "Fonds" in pf.columns:
            pf = pf[pf["Fonds"].isin(fonds_perimetre)]
        if nb is not None and hasattr(nb, "columns"):
            cols_k = [c for c in nb.columns if c in fonds_perimetre]
            nb = nb[cols_k] if cols_k else pd.DataFrame()

    nav_w_in = NAV_W / 72.0
    nav_h_in = NAV_H / 72.0
    buf_nav  = _make_nav_png(nb, nav_w_in, nav_h_in)
    img_nav  = Image(buf_nav, width=NAV_W, height=NAV_H)
    cap_nav  = Paragraph("Figure — Evolution NAV Base 100", styles["caption"])
    elements.append(KeepTogether([_center_image(img_nav, NAV_W, NAV_H, NAV_PAD), cap_nav]))
    elements.append(Spacer(1, 18))

    if pf.empty:
        elements.append(Paragraph("Aucune donnee de performance disponible.", styles["body"]))
        return elements

    elements.append(Paragraph("Tableau des Performances", styles["subsect"]))
    elements.append(ColorRect(USABLE_W, 1, COL_GRIS))
    elements.append(Spacer(1, 8))

    perf_cols = {"Perf 1M (%)", "Perf YTD (%)", "Perf Periode (%)"}
    col_order = ["Fonds", "NAV Derniere", "Base 100 Actuel",
                 "Perf 1M (%)", "Perf YTD (%)", "Perf Periode (%)"]
    available = [c for c in col_order if c in pf.columns]
    n = len(available)
    if n == 0:
        elements.append(Paragraph("Colonnes manquantes.", styles["body"]))
        return elements

    fonds_r    = 0.28
    rest_r     = (1.0 - fonds_r) / max(n - 1, 1)
    col_widths = [USABLE_W * (fonds_r if i == 0 else rest_r) for i in range(n)]
    hrow = [Paragraph(c, styles["pth_l"] if i == 0 else styles["pth_r"])
            for i, c in enumerate(available)]
    rows = [hrow]

    for _, row in pf.iterrows():
        data_row = []
        for i, col in enumerate(available):
            val = row.get(col)
            if col == "Fonds":
                data_row.append(Paragraph(str(val), styles["ptd_l"]))
            elif col in perf_cols:
                try:
                    fval = float(val)
                    valid = not np.isnan(fval)
                except (TypeError, ValueError):
                    valid = False
                    fval  = 0.0
                if not valid:
                    data_row.append(Paragraph("n.d.", styles["ptd"]))
                else:
                    sign = "+" if fval > 0 else ""
                    sty  = styles["ppos"] if fval >= 0 else styles["pneg"]
                    data_row.append(Paragraph("{}{:.2f}%".format(sign, fval), sty))
            else:
                try:
                    fval = float(val)
                    txt  = "{:.4f}".format(fval) if "NAV" in col else "{:.2f}".format(fval)
                except (TypeError, ValueError):
                    txt = "-"
                data_row.append(Paragraph(txt, styles["ptd"]))
        rows.append(data_row)

    ptbl = Table(rows, colWidths=col_widths, repeatRows=1)
    ptbl.setStyle(TableStyle([
        ("BACKGROUND",     (0, 0), (-1, 0),  COL_MARINE),
        ("LINEBELOW",      (0, 0), (-1, 0),  1.2, COL_ORANGE),
        ("GRID",           (0, 0), (-1, -1), 0.3, COL_GRIS),
        ("VALIGN",         (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",     (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING",  (0, 0), (-1, -1), 5),
        ("LEFTPADDING",    (0, 0), (-1, -1), 6),
        ("RIGHTPADDING",   (0, 0), (-1, -1), 6),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [COL_HEADER, COL_BLANC]),
    ]))
    elements.append(ptbl)
    return elements


# ---------------------------------------------------------------------------
# EN-TETE / PIED DE PAGE
# ---------------------------------------------------------------------------
def _header_footer(canvas, doc):
    canvas.saveState()
    if doc.page > 1:
        canvas.setFillColor(COL_MARINE)
        canvas.rect(0, 0, PAGE_W, 0.90 * cm, fill=1, stroke=0)
        canvas.setFont("Helvetica", 6.5)
        canvas.setFillColor(HexColor("#7ab8d8"))
        canvas.drawString(2 * cm, 0.30 * cm,
                          "Executive Report - Asset Management Division")
        canvas.drawRightString(PAGE_W - 2 * cm, 0.30 * cm,
                               "Page {}".format(doc.page))
        # Filet orange en haut de page (cohérent avec l'accent UI)
        canvas.setStrokeColor(COL_ORANGE)
        canvas.setLineWidth(1.8)
        canvas.line(0, PAGE_H - 0.38 * cm, PAGE_W, PAGE_H - 0.38 * cm)
        canvas.setFont("Helvetica", 6.5)
        canvas.setFillColor(COL_MARINE)
        canvas.drawRightString(PAGE_W - 2 * cm, PAGE_H - 0.26 * cm, "CONFIDENTIEL")
    canvas.restoreState()


# ---------------------------------------------------------------------------
# ENTREE PRINCIPALE
# ---------------------------------------------------------------------------
def generate_pdf(
    pipeline_df,
    kpis,
    aum_by_region=None,
    mode_comex=False,
    perf_data=None,
    nav_base100_df=None,
    fonds_perimetre=None,
    include_top10=True,
    include_outflows=False,
    include_perf=True,
):
    aum_by_region = aum_by_region or {}
    outflows      = kpis.get("outflows", [])

    if mode_comex:
        pipeline_df = pipeline_df.copy()
        pipeline_df["nom_client"] = pipeline_df.apply(
            lambda r: "{} - {}".format(r.get("type_client", ""), r.get("region", "")),
            axis=1
        )
        top_deals_safe = [
            dict(d, nom_client="{} - {}".format(d.get("type_client", ""), d.get("region", "")))
            for d in kpis.get("top_deals", [])
        ]
        outflows_safe = [
            dict(d, nom_client="{} - {}".format(d.get("type_client", ""), d.get("region", "")))
            for d in outflows
        ]
        kpis = dict(kpis, top_deals=top_deals_safe)
    else:
        top_deals_safe = kpis.get("top_deals", [])
        outflows_safe  = outflows

    styles  = _build_styles()
    pdf_buf = io.BytesIO()
    doc = SimpleDocTemplate(
        pdf_buf, pagesize=A4,
        leftMargin=MARGIN_H, rightMargin=MARGIN_H,
        topMargin=MARGIN_V,  bottomMargin=1.4 * cm,
        title="Executive Report - Asset Management",
        author="Asset Management Division",
    )

    elements = []
    elements += _page_garde(styles, mode_comex, kpis, fonds_perimetre)
    elements += _section_donuts(kpis, aum_by_region, styles)
    if include_top10:
        elements += _section_top10(top_deals_safe, outflows_safe, styles,
                                   mode_comex, include_outflows)
    elements += _section_pipeline(pipeline_df, styles, mode_comex)
    if include_perf and perf_data is not None and not perf_data.empty:
        elements += _section_performance(perf_data, nav_base100_df, styles, fonds_perimetre)

    doc.build(elements, onFirstPage=_header_footer, onLaterPages=_header_footer)
    result = pdf_buf.getvalue()
    pdf_buf.close()
    return result
