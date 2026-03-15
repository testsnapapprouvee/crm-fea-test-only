# =============================================================================
# pdf_generator.py  —  CRM Asset Management  —  Amundi Edition
# Design : reproduit fidèlement le modèle Canva fourni
#
# Page 1 (garde) : dessin direct canvas — fond marine plein, formes géométriques,
#                  textes positionnés en %, KPI strip, disclaimer, footer
# Pages suivantes : sections flowable standard avec en-tête/pied de page
# =============================================================================

import io
import math
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
from reportlab.lib.colors import HexColor, Color
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    Image, PageBreak, KeepTogether,
)
from reportlab.platypus.flowables import Flowable
from reportlab.pdfgen import canvas as rl_canvas


# ---------------------------------------------------------------------------
# COULEURS
# ---------------------------------------------------------------------------
COL_MARINE  = HexColor("#001c4b")
COL_CIEL    = HexColor("#019ee1")
COL_CIEL2   = HexColor("#089ee0")   # #089ee0 utilisé dans le modèle Canva
COL_ORANGE  = HexColor("#f07d00")
COL_BLANC   = HexColor("#ffffff")
COL_GRIS    = HexColor("#e8e8e8")
COL_TEXTE   = HexColor("#444444")
COL_BLEU_MD = HexColor("#7ab8d8")   # bleu moyen textes secondaires garde
COL_VERT    = HexColor("#1a7a3c")
COL_ROUGE   = HexColor("#8b2020")
COL_HEADER  = HexColor("#f0f5fa")
COL_NOIR    = HexColor("#000000")

HX_MARINE = "#001c4b"
HX_CIEL   = "#019ee1"
HX_BLANC  = "#ffffff"
HX_GRIS   = "#e8e8e8"
HX_TEXTE  = "#444444"

# ---------------------------------------------------------------------------
# PAGE
# ---------------------------------------------------------------------------
PAGE_W, PAGE_H = A4          # 595.3 x 841.9 pts
MARGIN_H  = 2.0 * cm
MARGIN_V  = 1.8 * cm
USABLE_W  = PAGE_W - 2 * MARGIN_H

DONUT_W   = USABLE_W * 0.46
DONUT_H   = DONUT_W  * 0.95
TOP10_W   = USABLE_W
NAV_W     = USABLE_W
NAV_H     = NAV_W * 0.40

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
        self.width = width; self.height = height; self.color = color
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


def _pct(x_pct, w=PAGE_W):
    """Convertit un % en pts (axe X)."""
    return w * x_pct / 100.0

def _pct_y(y_pct, h=PAGE_H):
    """
    Convertit un % du haut (Canva) en pts depuis le bas (ReportLab).
    Canva: 0% = haut, 100% = bas.
    ReportLab: 0 = bas, PAGE_H = haut.
    """
    return h * (1.0 - y_pct / 100.0)


# ---------------------------------------------------------------------------
# PAGE DE GARDE — dessin canvas direct (fidèle au modèle Canva)
# ---------------------------------------------------------------------------
def _draw_cover_page(c, mode_comex, kpis, fonds_perimetre):
    """
    Dessine la page de garde directement sur le canvas ReportLab.
    Reproduit le positionnement % du HTML Canva fourni.
    Coordonnées Canva converties : x% → _pct(x%), y% → _pct_y(y%)
    """
    W = PAGE_W
    H = PAGE_H

    # ---- Fond marine plein ----
    c.setFillColor(COL_MARINE)
    c.rect(0, 0, W, H, fill=1, stroke=0)

    # ---- Zone centrale claire (bande disclaimer) ----
    # top: 44.96%, height: 3.24% → en pts
    disc_top = _pct_y(44.96)          # haut de la bande
    disc_h   = H * 3.24 / 100.0
    disc_y   = disc_top - disc_h      # bas de la bande (ReportLab)
    c.setFillColor(HexColor("#e8e8e8"))
    c.rect(0, disc_y, W, disc_h, fill=1, stroke=0)

    # ---- Élément décoratif : grand arc bleu ciel (right side) ----
    # Cercle partiel en haut à droite pour rappeler le design Canva
    c.setFillColor(COL_CIEL)
    c.setStrokeColor(COL_CIEL)
    c.setLineWidth(0)
    # Grand arc décoratif — cercle centré hors page en haut à droite
    cx = W * 1.08
    cy = H * 0.78
    r  = H * 0.52
    c.saveState()
    path = c.beginPath()
    path.arc(cx - r, cy - r, cx + r, cy + r, startAng=100, extent=120)
    path.lineTo(cx, cy)
    path.close()
    c.setFillColor(COL_CIEL)
    c.drawPath(path, fill=1, stroke=0)
    c.restoreState()

    # Petit arc décoratif supplémentaire (forme Canva en bas à gauche)
    c.saveState()
    path2 = c.beginPath()
    cx2 = W * (-0.05)
    cy2 = H * 0.15
    r2  = H * 0.28
    path2.arc(cx2 - r2, cy2 - r2, cx2 + r2, cy2 + r2, startAng=290, extent=80)
    path2.lineTo(cx2, cy2)
    path2.close()
    c.setFillColor(HexColor("#012a6b"))  # marine légèrement plus clair
    c.drawPath(path2, fill=1, stroke=0)
    c.restoreState()

    # ---- Ligne décorative horizontale ciel sous le titre ----
    line_y = _pct_y(28.0)
    c.setStrokeColor(COL_CIEL)
    c.setLineWidth(1.2)
    c.line(_pct(9.15), line_y, _pct(55.0), line_y)

    # ---- Textes de la page de garde ----
    perim_str = ", ".join(fonds_perimetre) if fonds_perimetre else "Tous les fonds"
    today_str = date.today().strftime("%d %B %Y")

    # "INTERNAL PURPOSE ONLY" — top: 2.88%, left: 9.15%
    c.setFont("Helvetica-Bold", 8)
    c.setFillColor(COL_CIEL2)
    c.drawString(_pct(9.15), _pct_y(2.88) - 8, "INTERNAL PURPOSE ONLY")
    if mode_comex:
        c.drawString(_pct(9.15), _pct_y(2.88) - 18, "| MODE COMEX ACTIF — Noms anonymises")

    # "Executive Report" — top: 7.07%, bold, 28pt
    c.setFont("Helvetica-Bold", 28)
    c.setFillColor(COL_BLANC)
    c.drawString(_pct(9.15), _pct_y(7.07) - 28, "Executive Report")

    # "Pipeline & Reporting" — top: 10.20%, 18pt
    c.setFont("Helvetica", 18)
    c.setFillColor(COL_BLANC)
    c.drawString(_pct(9.15), _pct_y(10.20) - 18, "Pipeline & Reporting")

    # "Amundi Asset Management" — top: 15.08%, bleu moyen, 13pt
    c.setFont("Helvetica", 13)
    c.setFillColor(COL_BLEU_MD)
    c.drawString(_pct(9.15), _pct_y(15.08) - 13, "Amundi Asset Management")

    # Périmètre fonds — top: 17.22%, 13pt blanc
    c.setFont("Helvetica", 11)
    c.setFillColor(COL_BLANC)
    perim_display = "Perimetre : " + (perim_str[:80] + "..." if len(perim_str) > 80 else perim_str)
    c.drawString(_pct(9.15), _pct_y(17.22) - 11, perim_display)

    # Date — top: 23.53%, bleu ciel, 15pt
    c.setFont("Helvetica-Bold", 14)
    c.setFillColor(COL_CIEL2)
    c.drawString(_pct(9.15), _pct_y(23.53) - 14, today_str)

    # ---- Texte disclaimer (dans la bande grise) ----
    # top: 44.96%, centré
    disc_text_y = disc_y + disc_h * 0.35
    c.setFont("Helvetica-Oblique", 7)
    c.setFillColor(COL_NOIR)
    disc_line1 = ("Document strictement confidentiel a usage interne exclusif. "
                  "Reproduction et diffusion externe interdites.")
    disc_line2 = ("Les performances passees ne prejudgent pas des performances futures."
                  + (" Mode confidentiel actif : noms clients anonymises." if mode_comex else ""))
    c.drawCentredString(W / 2, disc_text_y + 8, disc_line1)
    c.drawCentredString(W / 2, disc_text_y - 2, disc_line2)

    # ---- Bande KPI (fond marine, après disclaimer) ----
    # Les KPIs Canva sont à top: ~35.5%  soit juste sous la bande grise
    kpi_band_top = disc_y - H * 0.01       # 1% sous la bande grise
    kpi_band_h   = H * 0.085               # hauteur bande KPI
    kpi_band_y   = kpi_band_top - kpi_band_h

    # Fond marine pour la bande KPI
    c.setFillColor(COL_MARINE)
    c.rect(0, kpi_band_y, W, kpi_band_h, fill=1, stroke=0)

    # Séparateurs verticaux ciel entre les KPIs
    kpi_items = [
        ("AUM Total",        fmt_aum(kpis.get("total_funded", 0))),
        ("Pipeline Actif",   fmt_aum(kpis.get("pipeline_actif", 0))),
        ("Taux Conversion",  "{:.1f}%".format(kpis.get("taux_conversion", 0))),
        ("Deals Actifs",     str(kpis.get("nb_deals_actifs", 0))),
    ]
    # Positions X Canva : 9.87%, 32.75%, 55.14%, 81.09%
    kpi_x_pcts = [9.87, 32.75, 55.14, 81.09]

    for i, ((label, value), x_pct) in enumerate(zip(kpi_items, kpi_x_pcts)):
        x = _pct(x_pct)
        label_y = kpi_band_y + kpi_band_h * 0.62
        value_y = kpi_band_y + kpi_band_h * 0.20

        # Label
        c.setFont("Helvetica", 7.5)
        c.setFillColor(COL_BLANC)
        c.drawCentredString(x + _pct(5), label_y, label)

        # Valeur (bold, plus grand)
        c.setFont("Helvetica-Bold", 13)
        c.setFillColor(COL_BLANC)
        c.drawString(x, value_y, value)

        # Séparateur vertical ciel (entre KPIs, pas après le dernier)
        if i < 3:
            sep_x = _pct(kpi_x_pcts[i+1]) - _pct(2.5)
            c.setStrokeColor(COL_CIEL)
            c.setLineWidth(0.8)
            c.line(sep_x, kpi_band_y + kpi_band_h * 0.15,
                   sep_x, kpi_band_y + kpi_band_h * 0.85)

    # ---- Footer page 1 ----
    footer_y = H * 0.025
    c.setFont("Helvetica", 7)
    c.setFillColor(COL_BLEU_MD)
    c.drawString(_pct(9.87), footer_y, "Executive Report - Internal Only")
    c.drawRightString(W - _pct(4), footer_y, "Page 1")


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
            at.set_fontsize(7.5); at.set_color(HX_BLANC); at.set_fontweight("bold")
        total = sum(values)
        ax.text(0, 0.10, fmt_aum(total), ha="center", va="center",
                fontsize=8, fontweight="bold", color=HX_MARINE)
        ax.text(0, -0.15, "Finance", ha="center", va="center",
                fontsize=6.5, color=HX_TEXTE)
        patches = [mpatches.Patch(color=colors[i],
                   label="{}: {}".format(labels[i], fmt_aum(values[i])))
                   for i in range(len(labels))]
        ax.legend(handles=patches, loc="lower center",
                  bbox_to_anchor=(0.5, -0.32), ncol=2,
                  fontsize=6.5, frameon=False, labelcolor=HX_MARINE)
        ax.set_title(title, fontsize=8.5, fontweight="bold",
                     color=HX_MARINE, pad=7)

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
        ax.text(0.5, 0.5, "Aucun deal", ha="center", va="center",
                fontsize=9, color=HX_MARINE, transform=ax.transAxes)
        ax.axis("off")
    else:
        labels = [
            "{} - {}".format(d.get("type_client",""), d.get("region",""))
            if mode_comex else str(d.get("nom_client",""))
            for d in deals_10
        ]
        values = [float(d.get("funded_aum", 0)) for d in deals_10]
        pairs  = sorted(zip(values, labels), reverse=True)
        values = [v for v, _ in pairs]
        labels = [l for _, l in pairs]
        max_v  = max(values) if values else 1.0

        colors = [HX_CIEL if i == 0 else HX_MARINE for i in range(len(deals_10))]
        bars   = ax.barh(range(len(labels)), values, color=colors,
                         edgecolor=HX_BLANC, height=0.65, linewidth=0.3)
        for bar, val in zip(bars, values):
            ax.text(bar.get_width() + max_v * 0.015,
                    bar.get_y() + bar.get_height() / 2,
                    fmt_aum(val), va="center", ha="left",
                    fontsize=8, color=HX_MARINE, fontweight="bold")

        ax.set_yticks(range(len(labels)))
        ax.set_yticklabels([l[:28] for l in labels], fontsize=8, color=HX_MARINE)
        ax.invert_yaxis()
        ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: fmt_aum(x)))
        ax.tick_params(axis="x", labelsize=7.5, colors=HX_MARINE)
        ax.set_xlim(0, max_v * 1.32)
        ax.grid(axis="x", alpha=0.22, color=HX_GRIS, linewidth=0.35)
        ax.spines["left"].set_color(HX_GRIS)
        ax.spines["bottom"].set_color(HX_GRIS)
        ax.set_title(title, fontsize=10, fontweight="bold", color=HX_MARINE, pad=10)

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
        NAV_COLORS = [HX_MARINE, HX_CIEL, "#1a5e8a", "#4a8fbd",
                      "#003f7a", "#2c7fb8", "#004f8c", "#6baed6"]
        for i, col in enumerate(nav_df.columns):
            series = nav_df[col].dropna()
            if series.empty: continue
            color = NAV_COLORS[i % len(NAV_COLORS)]
            if len(series) >= 2:
                ax.plot(series.index, series.values, label=col,
                        color=color, linewidth=1.5, alpha=0.92)
            else:
                ax.scatter(series.index, series.values, color=color, s=50, label=col)

        ax.axhline(100, color=HX_GRIS, linewidth=0.7, linestyle="dotted")
        ax.set_ylabel("Base 100", fontsize=7.5, color=HX_MARINE)
        ax.tick_params(colors=HX_MARINE, labelsize=7)
        ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)
        ax.spines["left"].set_color(HX_GRIS); ax.spines["bottom"].set_color(HX_GRIS)
        ax.grid(axis="y", alpha=0.18, color=HX_GRIS, linewidth=0.35)
        ax.grid(axis="x", visible=False)
        ax.legend(fontsize=7, frameon=True, framealpha=0.90,
                  edgecolor=HX_GRIS, labelcolor=HX_MARINE, loc="upper left")
        ax.set_title("Evolution NAV - Base 100", fontsize=9,
                     fontweight="bold", color=HX_MARINE, pad=8)
        plt.xticks(rotation=12, ha="right", fontsize=7)

    fig.subplots_adjust(left=0.08, right=0.97, top=0.90, bottom=0.14)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, facecolor=HX_BLANC)
    plt.close(fig)
    buf.seek(0)
    return buf


# ---------------------------------------------------------------------------
# STYLES REPORTLAB
# ---------------------------------------------------------------------------
def _build_styles():
    s_cover   = ParagraphStyle("cover",   fontName="Helvetica",         fontSize=9,   textColor=COL_BLANC,  leading=17)
    s_section = ParagraphStyle("section", fontName="Helvetica-Bold",    fontSize=11,  textColor=COL_CIEL,   spaceBefore=10, spaceAfter=4)
    s_subsect = ParagraphStyle("subsect", fontName="Helvetica-Bold",    fontSize=9,   textColor=COL_MARINE, spaceBefore=6,  spaceAfter=3)
    s_body    = ParagraphStyle("body",    fontName="Helvetica",         fontSize=8.5, textColor=COL_TEXTE,  spaceAfter=4, leading=13)
    s_th      = ParagraphStyle("th",      fontName="Helvetica-Bold",    fontSize=7.5, textColor=COL_BLANC,  alignment=TA_LEFT)
    s_th_r    = ParagraphStyle("th_r",    fontName="Helvetica-Bold",    fontSize=7.5, textColor=COL_BLANC,  alignment=TA_RIGHT)
    s_td      = ParagraphStyle("td",      fontName="Helvetica",         fontSize=7.5, textColor=COL_MARINE, alignment=TA_LEFT)
    s_td_r    = ParagraphStyle("td_r",    fontName="Helvetica",         fontSize=7.5, textColor=COL_MARINE, alignment=TA_RIGHT)
    s_td_g    = ParagraphStyle("td_g",    fontName="Helvetica",         fontSize=7.5, textColor=COL_TEXTE,  alignment=TA_LEFT)
    s_kpi_lbl = ParagraphStyle("kpi_lbl", fontName="Helvetica",         fontSize=7,   textColor=COL_BLANC,  alignment=TA_CENTER)
    s_kpi_val = ParagraphStyle("kpi_val", fontName="Helvetica-Bold",    fontSize=15,  textColor=COL_BLANC,  alignment=TA_CENTER)
    s_disc    = ParagraphStyle("disc",    fontName="Helvetica-Oblique", fontSize=6.5, textColor=COL_TEXTE,  alignment=TA_CENTER, leading=10)
    s_pth_l   = ParagraphStyle("pth_l",  fontName="Helvetica-Bold",    fontSize=7.5, textColor=COL_BLANC,  alignment=TA_LEFT)
    s_pth_r   = ParagraphStyle("pth_r",  fontName="Helvetica-Bold",    fontSize=7.5, textColor=COL_BLANC,  alignment=TA_RIGHT)
    s_ptd_l   = ParagraphStyle("ptd_l",  fontName="Helvetica",         fontSize=7.5, textColor=COL_MARINE, alignment=TA_LEFT)
    s_ptd     = ParagraphStyle("ptd",    fontName="Helvetica",         fontSize=7.5, textColor=COL_MARINE, alignment=TA_RIGHT)
    s_ppos    = ParagraphStyle("ppos",   fontName="Helvetica-Bold",    fontSize=7.5, textColor=COL_VERT,   alignment=TA_RIGHT)
    s_pneg    = ParagraphStyle("pneg",   fontName="Helvetica-Bold",    fontSize=7.5, textColor=COL_ROUGE,  alignment=TA_RIGHT)
    s_caption = ParagraphStyle("caption",fontName="Helvetica-Oblique", fontSize=7,   textColor=COL_TEXTE,  alignment=TA_CENTER, spaceAfter=3)
    s_alert   = ParagraphStyle("alert",  fontName="Helvetica-Bold",    fontSize=7.5, textColor=COL_ORANGE, alignment=TA_LEFT)
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
# PAGE 2 : DONUTS
# ---------------------------------------------------------------------------
def _section_donuts(kpis, aum_by_region, styles):
    elements = []
    elements.append(Paragraph("Repartition AUM Finance", styles["section"]))
    elements.append(ColorRect(USABLE_W, 2, COL_CIEL))
    elements.append(Spacer(1, 14))

    d_w_in = DONUT_W / 72.0
    d_h_in = DONUT_H / 72.0

    abt    = kpis.get("aum_by_type", {})
    buf_d1 = _make_donut_png(list(abt.keys()), list(abt.values()),
                             "AUM par Type de Client", d_w_in, d_h_in)
    img_d1 = Image(buf_d1, width=DONUT_W, height=DONUT_H)

    buf_d2 = _make_donut_png(list(aum_by_region.keys()), list(aum_by_region.values()),
                             "AUM par Region Geographique", d_w_in, d_h_in)
    img_d2 = Image(buf_d2, width=DONUT_W, height=DONUT_H)

    gap = USABLE_W - 2 * DONUT_W
    tbl = Table([[img_d1, Spacer(gap, 1), img_d2]],
                colWidths=[DONUT_W, gap, DONUT_W])
    tbl.setStyle(TableStyle([
        ("ALIGN", (0,0),(-1,-1), "CENTER"),
        ("VALIGN",(0,0),(-1,-1), "TOP"),
        ("LEFTPADDING",  (0,0),(-1,-1), 0),
        ("RIGHTPADDING", (0,0),(-1,-1), 0),
        ("TOPPADDING",   (0,0),(-1,-1), 0),
        ("BOTTOMPADDING",(0,0),(-1,-1), 0),
    ]))
    cap = Paragraph(
        "Figure 1 — Repartition AUM Finance par Type de Client et par Region",
        styles["caption"])
    elements.append(KeepTogether([tbl, cap]))
    return elements


# ---------------------------------------------------------------------------
# PAGE 3 : TOP 10
# ---------------------------------------------------------------------------
def _section_top10(top_deals, outflows, styles, mode_comex, include_outflows):
    elements = []
    elements.append(PageBreak())
    elements.append(Paragraph("Top 10 Deals", styles["section"]))
    elements.append(ColorRect(USABLE_W, 2, COL_CIEL))
    elements.append(Spacer(1, 16))

    elements.append(Paragraph("Inflows — AUM Finance (Funded)", styles["subsect"]))
    elements.append(Spacer(1, 8))

    n_in      = max(len(top_deals[:10]), 1)
    t10_h_pts = max(180, min(320, n_in * 27 + 40))
    t10_w_in  = TOP10_W / 72.0
    t10_h_in  = t10_h_pts / 72.0

    buf_in = _make_top10_png(top_deals, mode_comex, t10_w_in, t10_h_in,
                             title="Top 10 Inflows — AUM Finance")
    img_in = Image(buf_in, width=TOP10_W, height=t10_h_pts)
    cap_in = Paragraph("Figure 2 — Top 10 Deals par AUM Finance (statut Funded)",
                       styles["caption"])
    elements.append(KeepTogether([img_in, cap_in]))

    if include_outflows and outflows:
        elements.append(Spacer(1, 22))
        elements.append(Paragraph("Outflows — AUM Rachete (Redeemed)", styles["subsect"]))
        elements.append(Spacer(1, 8))
        n_out     = max(len(outflows[:10]), 1)
        out_h_pts = max(120, min(260, n_out * 27 + 40))
        out_h_in  = out_h_pts / 72.0
        buf_out = _make_top10_png(outflows, mode_comex, t10_w_in, out_h_in,
                                  title="Top 10 Outflows — AUM Rachete")
        img_out = Image(buf_out, width=TOP10_W, height=out_h_pts)
        cap_out = Paragraph("Figure 3 — Top 10 Rachats (statut Redeemed)",
                            styles["caption"])
        elements.append(KeepTogether([img_out, cap_out]))

    return elements


# ---------------------------------------------------------------------------
# PAGE 4 : PIPELINE
# ---------------------------------------------------------------------------
def _section_pipeline(pipeline_df, styles, mode_comex):
    elements = []
    elements.append(PageBreak())
    elements.append(Paragraph("Pipeline Actif - Recapitulatif", styles["section"]))
    elements.append(ColorRect(USABLE_W, 2, COL_CIEL))
    elements.append(Spacer(1, 10))

    actifs = pipeline_df[pipeline_df["statut"].isin(
        ["Prospect", "Initial Pitch", "Due Diligence", "Soft Commit"]
    )].copy()

    if actifs.empty:
        elements.append(Paragraph("Aucun deal actif dans le perimetre.", styles["body"]))
    else:
        if mode_comex:
            actifs["nom_client"] = actifs.apply(
                lambda r: "{} - {}".format(r.get("type_client",""), r.get("region","")), axis=1)
        ratios     = [0.22, 0.16, 0.12, 0.13, 0.13, 0.13, 0.11]
        col_widths = [USABLE_W * r for r in ratios]
        headers    = ["Client", "Fonds", "Statut", "AUM Cible",
                      "AUM Revise", "Prochaine Action", "Commercial"]
        rows       = [[Paragraph(h, styles["th"]) for h in headers]]
        today      = date.today()
        for _, row in actifs.iterrows():
            nad = row.get("next_action_date")
            if isinstance(nad, date):
                nad_str   = "[!] {}".format(nad.isoformat()) if nad < today else nad.isoformat()
                nad_style = styles["alert"] if nad < today else styles["td"]
            else:
                nad_str = "-"; nad_style = styles["td"]
            rows.append([
                Paragraph(str(row.get("nom_client",""))[:26], styles["td"]),
                Paragraph(str(row.get("fonds","")),            styles["td"]),
                Paragraph(str(row.get("statut","")),           styles["td"]),
                Paragraph(fmt_aum(float(row.get("target_aum_initial",0) or 0)), styles["td_r"]),
                Paragraph(fmt_aum(float(row.get("revised_aum",0) or 0)),         styles["td_r"]),
                Paragraph(nad_str, nad_style),
                Paragraph(str(row.get("sales_owner",""))[:14], styles["td"]),
            ])
        tbl = Table(rows, colWidths=col_widths, repeatRows=1)
        tbl.setStyle(TableStyle([
            ("BACKGROUND",     (0,0),(-1,0),  COL_MARINE),
            ("LINEBELOW",      (0,0),(-1,0),  1.5, COL_CIEL),
            ("GRID",           (0,0),(-1,-1), 0.3, COL_GRIS),
            ("VALIGN",         (0,0),(-1,-1), "MIDDLE"),
            ("TOPPADDING",     (0,0),(-1,-1), 4),
            ("BOTTOMPADDING",  (0,0),(-1,-1), 4),
            ("LEFTPADDING",    (0,0),(-1,-1), 5),
            ("RIGHTPADDING",   (0,0),(-1,-1), 5),
            ("ROWBACKGROUNDS", (0,1),(-1,-1), [COL_HEADER, COL_BLANC]),
        ]))
        elements.append(tbl)

    lost_paused = pipeline_df[pipeline_df["statut"].isin(["Lost","Paused"])].copy()
    if not lost_paused.empty:
        elements.append(Spacer(1, 16))
        elements.append(Paragraph("Deals Perdus / En Pause", styles["subsect"]))
        elements.append(ColorRect(USABLE_W, 1, COL_GRIS))
        elements.append(Spacer(1, 8))
        if mode_comex:
            lost_paused["nom_client"] = lost_paused.apply(
                lambda r: "{} - {}".format(r.get("type_client",""), r.get("region","")), axis=1)
        lp_ratios  = [0.26, 0.18, 0.12, 0.22, 0.22]
        lp_col_w   = [USABLE_W * r for r in lp_ratios]
        lp_rows    = [[Paragraph(h, styles["th"])
                       for h in ["Client","Fonds","Statut","Raison","Concurrent"]]]
        for _, row in lost_paused.iterrows():
            lp_rows.append([
                Paragraph(str(row.get("nom_client",""))[:24],          styles["td_g"]),
                Paragraph(str(row.get("fonds","")),                    styles["td_g"]),
                Paragraph(str(row.get("statut","")),                   styles["td_g"]),
                Paragraph(str(row.get("raison_perte","") or "-"),      styles["td_g"]),
                Paragraph(str(row.get("concurrent_choisi","") or "-"), styles["td_g"]),
            ])
        lp_tbl = Table(lp_rows, colWidths=lp_col_w, repeatRows=1)
        lp_tbl.setStyle(TableStyle([
            ("BACKGROUND",    (0,0),(-1,0),  HexColor("#7a7a7a")),
            ("GRID",          (0,0),(-1,-1), 0.3, COL_GRIS),
            ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
            ("TOPPADDING",    (0,0),(-1,-1), 4),
            ("BOTTOMPADDING", (0,0),(-1,-1), 4),
            ("LEFTPADDING",   (0,0),(-1,-1), 5),
            ("RIGHTPADDING",  (0,0),(-1,-1), 5),
        ]))
        elements.append(lp_tbl)
    return elements


# ---------------------------------------------------------------------------
# PAGE 5 : PERFORMANCE NAV
# ---------------------------------------------------------------------------
def _section_performance(perf_data, nav_base100_df, styles, fonds_perimetre):
    elements = []
    elements.append(PageBreak())
    elements.append(Paragraph("Analyse de Performance - NAV", styles["section"]))
    elements.append(ColorRect(USABLE_W, 2, COL_CIEL))
    elements.append(Spacer(1, 14))

    pf = perf_data.copy() if perf_data is not None else pd.DataFrame()
    nb = nav_base100_df.copy() if nav_base100_df is not None else None

    if fonds_perimetre:
        if not pf.empty and "Fonds" in pf.columns:
            pf = pf[pf["Fonds"].isin(fonds_perimetre)]
        if nb is not None and hasattr(nb, "columns"):
            cols_k = [c for c in nb.columns if c in fonds_perimetre]
            nb = nb[cols_k] if cols_k else pd.DataFrame()

    nav_w_in = NAV_W / 72.0; nav_h_in = NAV_H / 72.0
    buf_nav  = _make_nav_png(nb, nav_w_in, nav_h_in)
    img_nav  = Image(buf_nav, width=NAV_W, height=NAV_H)
    tbl_nav  = Table([[img_nav]], colWidths=[NAV_W])
    tbl_nav.setStyle(TableStyle([("ALIGN",(0,0),(-1,-1),"CENTER"),
                                  ("LEFTPADDING",(0,0),(-1,-1),0),
                                  ("RIGHTPADDING",(0,0),(-1,-1),0),
                                  ("TOPPADDING",(0,0),(-1,-1),0),
                                  ("BOTTOMPADDING",(0,0),(-1,-1),0)]))
    cap_nav = Paragraph("Figure — Evolution NAV Base 100", styles["caption"])
    elements.append(KeepTogether([tbl_nav, cap_nav]))
    elements.append(Spacer(1, 16))

    if pf.empty:
        elements.append(Paragraph("Aucune donnee de performance.", styles["body"]))
        return elements

    elements.append(Paragraph("Tableau des Performances", styles["subsect"]))
    elements.append(ColorRect(USABLE_W, 1, COL_GRIS))
    elements.append(Spacer(1, 8))

    perf_cols = {"Perf 1M (%)", "Perf YTD (%)", "Perf Periode (%)"}
    col_order = ["Fonds","NAV Derniere","Base 100 Actuel",
                 "Perf 1M (%)","Perf YTD (%)","Perf Periode (%)"]
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
                    fval = float(val); valid = not np.isnan(fval)
                except (TypeError, ValueError):
                    valid = False; fval = 0.0
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
        ("BACKGROUND",     (0,0),(-1,0),  COL_MARINE),
        ("LINEBELOW",      (0,0),(-1,0),  1.2, COL_CIEL),
        ("GRID",           (0,0),(-1,-1), 0.3, COL_GRIS),
        ("VALIGN",         (0,0),(-1,-1), "MIDDLE"),
        ("TOPPADDING",     (0,0),(-1,-1), 5),
        ("BOTTOMPADDING",  (0,0),(-1,-1), 5),
        ("LEFTPADDING",    (0,0),(-1,-1), 6),
        ("RIGHTPADDING",   (0,0),(-1,-1), 6),
        ("ROWBACKGROUNDS", (0,1),(-1,-1), [COL_HEADER, COL_BLANC]),
    ]))
    elements.append(ptbl)
    return elements


# ---------------------------------------------------------------------------
# EN-TÊTE / PIED DE PAGE (pages 2+)
# Reproduit le style du modèle Canva : footer bas marine, ligne ciel en haut
# ---------------------------------------------------------------------------
def _make_header_footer_fn(mode_comex):
    """Retourne la fonction onPage avec mode_comex capturé."""
    def _header_footer(canvas, doc):
        canvas.saveState()
        if doc.page == 1:
            # Page de garde dessinée ailleurs
            canvas.restoreState()
            return

        W = PAGE_W; H = PAGE_H

        # Ligne ciel en haut
        canvas.setStrokeColor(COL_CIEL)
        canvas.setLineWidth(1.8)
        canvas.line(0, H - 0.38 * cm, W, H - 0.38 * cm)
        # "CONFIDENTIEL" en haut à droite
        canvas.setFont("Helvetica", 6.5)
        canvas.setFillColor(COL_MARINE)
        canvas.drawRightString(W - 2*cm, H - 0.26*cm, "CONFIDENTIEL")

        # Bande footer marine
        footer_h = 0.90 * cm
        canvas.setFillColor(COL_MARINE)
        canvas.rect(0, 0, W, footer_h, fill=1, stroke=0)
        # Textes footer
        canvas.setFont("Helvetica", 6.5)
        canvas.setFillColor(COL_BLEU_MD)
        canvas.drawString(2*cm, 0.30*cm, "Executive Report - Internal Only")
        canvas.drawRightString(W - 2*cm, 0.30*cm, "Page {}".format(doc.page))

        canvas.restoreState()

    return _header_footer


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
            lambda r: "{} - {}".format(r.get("type_client",""), r.get("region","")), axis=1)
        top_deals_safe = [
            dict(d, nom_client="{} - {}".format(d.get("type_client",""), d.get("region","")))
            for d in kpis.get("top_deals", [])
        ]
        outflows_safe = [
            dict(d, nom_client="{} - {}".format(d.get("type_client",""), d.get("region","")))
            for d in outflows
        ]
        kpis = dict(kpis, top_deals=top_deals_safe)
    else:
        top_deals_safe = kpis.get("top_deals", [])
        outflows_safe  = outflows

    styles   = _build_styles()
    pdf_buf  = io.BytesIO()
    hf_fn    = _make_header_footer_fn(mode_comex)

    # Marge page 1 = 0 (dessin full-bleed), pages suivantes = marges normales
    doc = SimpleDocTemplate(
        pdf_buf, pagesize=A4,
        leftMargin=0, rightMargin=0,
        topMargin=0,  bottomMargin=0,
        title="Executive Report - Asset Management",
        author="Asset Management Division",
    )

    # Flowables : page 1 = Spacer pleine page (le dessin est dans onFirstPage)
    # Pages 2+ = contenu normal avec marges
    elements = [Spacer(PAGE_W, PAGE_H)]   # page 1 vide (dessin direct)
    elements.append(PageBreak())

    # Wrapper pour remettre les marges à partir de la page 2
    # On utilise un SimpleDocTemplate séparé pour les pages 2+
    # et on fusionne les deux PDFs
    pdf_buf_content = io.BytesIO()
    doc_content = SimpleDocTemplate(
        pdf_buf_content, pagesize=A4,
        leftMargin=MARGIN_H, rightMargin=MARGIN_H,
        topMargin=MARGIN_V,  bottomMargin=1.4*cm,
        title="Executive Report", author="AM Division",
    )
    content_elems = []
    content_elems += _section_donuts(kpis, aum_by_region, styles)
    if include_top10:
        content_elems += _section_top10(top_deals_safe, outflows_safe, styles,
                                        mode_comex, include_outflows)
    content_elems += _section_pipeline(pipeline_df, styles, mode_comex)
    if include_perf and perf_data is not None and hasattr(perf_data,"empty") and not perf_data.empty:
        content_elems += _section_performance(perf_data, nav_base100_df, styles, fonds_perimetre)

    doc_content.build(content_elems, onFirstPage=hf_fn, onLaterPages=hf_fn)

    # Assembler : page de garde (canvas) + contenu
    # On utilise PdfReader/PdfWriter si PyPDF2 dispo, sinon on dessine en une passe
    try:
        from pypdf import PdfReader, PdfWriter
        _merge_with_pypdf(pdf_buf, pdf_buf_content, mode_comex, kpis, fonds_perimetre, hf_fn)
    except ImportError:
        try:
            from PyPDF2 import PdfReader, PdfWriter
            _merge_with_pypdf(pdf_buf, pdf_buf_content, mode_comex, kpis, fonds_perimetre, hf_fn)
        except ImportError:
            # Fallback : page de garde dans le même doc, dessin direct dans onFirstPage
            _build_single_pass(pdf_buf, mode_comex, kpis, fonds_perimetre,
                               pipeline_df, aum_by_region, top_deals_safe, outflows_safe,
                               styles, include_top10, include_outflows,
                               perf_data, nav_base100_df, include_perf)
            result = pdf_buf.getvalue()
            pdf_buf.close()
            pdf_buf_content.close()
            return result

    result = pdf_buf.getvalue()
    pdf_buf.close()
    pdf_buf_content.close()
    return result


def _merge_with_pypdf(pdf_buf, pdf_buf_content, mode_comex, kpis, fonds_perimetre, hf_fn):
    """Fusionne la page de garde (canvas seul) avec les pages de contenu."""
    try:
        from pypdf import PdfReader, PdfWriter
    except ImportError:
        from PyPDF2 import PdfReader, PdfWriter

    # Dessiner la page de garde seule
    cover_buf = io.BytesIO()
    c = rl_canvas.Canvas(cover_buf, pagesize=A4)
    _draw_cover_page(c, mode_comex, kpis, fonds_perimetre)
    c.showPage()
    c.save()
    cover_buf.seek(0)

    # Fusionner
    writer = PdfWriter()
    cover_reader = PdfReader(cover_buf)
    writer.add_page(cover_reader.pages[0])

    pdf_buf_content.seek(0)
    content_reader = PdfReader(pdf_buf_content)
    for page in content_reader.pages:
        writer.add_page(page)

    writer.write(pdf_buf)


def _build_single_pass(pdf_buf, mode_comex, kpis, fonds_perimetre,
                       pipeline_df, aum_by_region, top_deals_safe, outflows_safe,
                       styles, include_top10, include_outflows,
                       perf_data, nav_base100_df, include_perf):
    """Fallback sans pypdf : page de garde dessinée via onFirstPage."""
    def on_first(canvas, doc):
        _draw_cover_page(canvas, mode_comex, kpis, fonds_perimetre)

    hf_fn = _make_header_footer_fn(mode_comex)

    def on_later(canvas, doc):
        hf_fn(canvas, doc)

    doc = SimpleDocTemplate(
        pdf_buf, pagesize=A4,
        leftMargin=MARGIN_H, rightMargin=MARGIN_H,
        topMargin=MARGIN_V,  bottomMargin=1.4*cm,
        title="Executive Report", author="AM Division",
    )
    elements = [Spacer(1, PAGE_H), PageBreak()]
    elements += _section_donuts(kpis, aum_by_region, styles)
    if include_top10:
        elements += _section_top10(top_deals_safe, outflows_safe, styles,
                                   mode_comex, include_outflows)
    elements += _section_pipeline(pipeline_df, styles, mode_comex)
    if include_perf and perf_data is not None and hasattr(perf_data,"empty") and not perf_data.empty:
        elements += _section_performance(perf_data, nav_base100_df, styles, fonds_perimetre)

    doc.build(elements, onFirstPage=on_first, onLaterPages=on_later)
