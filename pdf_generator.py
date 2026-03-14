# =============================================================================
# pdf_generator.py  —  CRM Asset Management  —  Amundi Edition
# Priorite : STABILITE ABSOLUE + MISE EN PAGE CORRECTE
# Charte : #001c4b Marine | #019ee1 Ciel (dominant) | #f07d00 Orange (unique)
# Layout graphiques :
#   - Donuts Type + Region : COTE A COTE sur la meme ligne (Table 2 colonnes)
#   - Top 10 Deals         : dessous, largeur 85% centree
#   - Aucun chart ne deborde la page
# Matplotlib Agg UNIQUEMENT — zero Plotly dans ce fichier
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
# COULEURS — charte Amundi
# ---------------------------------------------------------------------------

COL_MARINE  = HexColor("#001c4b")
COL_CIEL    = HexColor("#019ee1")
COL_ORANGE  = HexColor("#f07d00")   # usage unique : badge RETARD dans tableau
COL_BLANC   = HexColor("#ffffff")
COL_GRIS    = HexColor("#e8e8e8")
COL_TEXTE   = HexColor("#444444")
COL_VERT    = HexColor("#1a7a3c")
COL_ROUGE   = HexColor("#8b2020")
COL_HEADER  = HexColor("#f0f5fa")   # fond alternatif lignes tableau

HX_MARINE   = "#001c4b"
HX_CIEL     = "#019ee1"
HX_BLANC    = "#ffffff"
HX_GRIS     = "#e8e8e8"
HX_TEXTE    = "#444444"

# ---------------------------------------------------------------------------
# MISE EN PAGE
# ---------------------------------------------------------------------------

PAGE_W, PAGE_H = A4           # 595.3 x 841.9 pts
MARGIN_H  = 2.0 * cm
MARGIN_V  = 1.8 * cm
USABLE_W  = PAGE_W - 2 * MARGIN_H    # 481.9 pts

# Donuts cote a cote — chacun occupe ~46% de la largeur utile
DONUT_W   = USABLE_W * 0.46          # 221.7 pts  ~7.8 cm
DONUT_H   = DONUT_W  * 0.95          # presque carre, legende en bas

# Top 10 — barre horizontale, 85% de la largeur, centre
TOP10_W   = USABLE_W * 0.85          # 409.6 pts ~14.5 cm
PAD_TOP10 = (USABLE_W - TOP10_W) / 2.0

# NAV Base 100 — pleine largeur utile, rapport 16/7
NAV_W     = USABLE_W
NAV_H     = NAV_W * 0.38

# Palette graphiques Matplotlib — bleu marine + variations
PALETTE = [
    "#1a5e8a", "#001c4b", "#4a8fbd", "#003f7a",
    "#2c7fb8", "#004f8c", "#6baed6", "#08519c",
    "#9ecae1", "#003060",
]

# rcParams — appliques avant chaque figure, plats
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
# FLOWABLE UTILITAIRE
# ---------------------------------------------------------------------------

class ColorRect(Flowable):
    """Filet de couleur horizontal."""
    def __init__(self, width, height, color):
        super().__init__()
        self.width  = width
        self.height = height
        self.color  = color

    def draw(self):
        self.canv.setFillColor(self.color)
        self.canv.rect(0, 0, self.width, self.height, fill=1, stroke=0)


# ---------------------------------------------------------------------------
# FORMATAGE
# ---------------------------------------------------------------------------

def fmt_aum(v):
    """M EUR / Md EUR avec 1 decimale."""
    try:
        fv = float(v)
    except (TypeError, ValueError):
        return "-"
    if fv >= 1_000_000_000:
        return "{:.1f} Md EUR".format(fv / 1_000_000_000)
    if fv >= 1_000_000:
        return "{:.1f} M EUR".format(fv / 1_000_000)
    if fv >= 1_000:
        return "{:.0f} k EUR".format(fv / 1_000)
    return "{:.0f} EUR".format(fv)


# ---------------------------------------------------------------------------
# GRAPHIQUES MATPLOTLIB
# Chaque fonction :
#   1. plt.rcParams.update(MPL_RC)
#   2. fig, ax = plt.subplots(figsize=(...))
#   3. figure content
#   4. fig.savefig(buf, format='png', dpi=150, facecolor=HX_BLANC)
#      NOTE: PAS de bbox_inches='tight' — on controle les dimensions nous-memes
#   5. plt.close(fig)
# ---------------------------------------------------------------------------

def _make_donut_png(labels, values, title, fig_w_in, fig_h_in):
    """
    Donut chart — dimensions en pouces passees explicitement.
    Retourne io.BytesIO PNG.
    bbox_inches NON utilise pour garder les dimensions exactes.
    """
    plt.rcParams.update(MPL_RC)
    fig, ax = plt.subplots(figsize=(fig_w_in, fig_h_in))
    fig.patch.set_facecolor(HX_BLANC)

    if not labels or sum(values) == 0:
        ax.text(0.5, 0.5, "Aucune donnee",
                ha="center", va="center", fontsize=9, color=HX_MARINE,
                transform=ax.transAxes)
        ax.axis("off")
    else:
        colors = [PALETTE[i % len(PALETTE)] for i in range(len(labels))]
        _, _, autotexts = ax.pie(
            values,
            colors=colors,
            autopct="%1.1f%%",
            startangle=90,
            pctdistance=0.72,
            wedgeprops={"width": 0.50, "edgecolor": HX_BLANC, "linewidth": 1.6},
        )
        for at in autotexts:
            at.set_fontsize(8)
            at.set_color(HX_BLANC)
            at.set_fontweight("bold")

        total = sum(values)
        ax.text(0,  0.10, fmt_aum(total),
                ha="center", va="center", fontsize=8.5,
                fontweight="bold", color=HX_MARINE)
        ax.text(0, -0.15, "Finance",
                ha="center", va="center", fontsize=6.5, color=HX_TEXTE)

        patches = [
            mpatches.Patch(color=colors[i],
                           label="{}: {}".format(labels[i][:14], fmt_aum(values[i])))
            for i in range(len(labels))
        ]
        ax.legend(
            handles=patches,
            loc="lower center",
            bbox_to_anchor=(0.5, -0.26),
            ncol=2,
            fontsize=6.5,
            frameon=False,
            labelcolor=HX_MARINE,
        )
        ax.set_title(title, fontsize=8.5, fontweight="bold",
                     color=HX_MARINE, pad=7)

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, facecolor=HX_BLANC)
    plt.close(fig)
    buf.seek(0)
    return buf


def _make_top10_png(top_deals, mode_comex, fig_w_in, fig_h_in):
    """
    Barre horizontale Top 10 — dimensions en pouces.
    Retourne io.BytesIO PNG.
    """
    plt.rcParams.update(MPL_RC)
    fig, ax = plt.subplots(figsize=(fig_w_in, fig_h_in))
    fig.patch.set_facecolor(HX_BLANC)
    ax.set_facecolor(HX_BLANC)

    deals = top_deals[:10]
    if not deals:
        ax.text(0.5, 0.5, "Aucun deal Funded",
                ha="center", va="center", fontsize=9, color=HX_MARINE,
                transform=ax.transAxes)
        ax.axis("off")
    else:
        labels = [
            "{} - {}".format(d.get("type_client", ""), d.get("region", ""))
            if mode_comex else str(d.get("nom_client", ""))
            for d in deals
        ]
        values = [float(d.get("funded_aum", 0)) for d in deals]
        pairs  = sorted(zip(values, labels), reverse=True)
        values = [v for v, _ in pairs]
        labels = [l for _, l in pairs]
        max_v  = max(values) if values else 1.0

        colors = [PALETTE[i % len(PALETTE)] for i in range(len(deals))]
        bars   = ax.barh(range(len(labels)), values,
                         color=colors, edgecolor=HX_BLANC,
                         height=0.60, linewidth=0.4)

        for bar, val in zip(bars, values):
            ax.text(
                bar.get_width() + max_v * 0.01,
                bar.get_y() + bar.get_height() / 2,
                fmt_aum(val),
                va="center", ha="left",
                fontsize=7, color=HX_MARINE, fontweight="bold",
            )

        ax.set_yticks(range(len(labels)))
        ax.set_yticklabels([l[:22] for l in labels], fontsize=7.5, color=HX_MARINE)
        ax.invert_yaxis()
        ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: fmt_aum(x)))
        ax.tick_params(axis="x", labelsize=7, colors=HX_MARINE)
        ax.set_xlim(0, max_v * 1.30)
        ax.grid(axis="x", alpha=0.20, color=HX_GRIS, linewidth=0.35)
        ax.spines["left"].set_color(HX_GRIS)
        ax.spines["bottom"].set_color(HX_GRIS)
        ax.set_title("Top 10 Deals - AUM Finance",
                     fontsize=9, fontweight="bold", color=HX_MARINE, pad=8)

    fig.subplots_adjust(left=0.22, right=0.88, top=0.92, bottom=0.08)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, facecolor=HX_BLANC)
    plt.close(fig)
    buf.seek(0)
    return buf


def _make_nav_png(nav_df, fig_w_in, fig_h_in):
    """
    Courbe NAV Base 100 — dimensions en pouces.
    Retourne io.BytesIO PNG.
    """
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
        # Palette NAV — couleurs vraiment distinctives
        NAV_COLORS = [
            "#001c4b", "#019ee1", "#e63946", "#2a9d8f",
            "#f4a261", "#9b59b6", "#27ae60", "#e67e22",
        ]
        plotted = 0
        for i, col in enumerate(nav_df.columns):
            series = nav_df[col].dropna()
            if series.empty:
                continue
            color = NAV_COLORS[i % len(NAV_COLORS)]
            if len(series) >= 2:
                ax.plot(series.index, series.values,
                        label=col, color=color,
                        linewidth=1.5, linestyle="solid", alpha=0.92)
            else:
                ax.scatter(series.index, series.values,
                           color=color, s=50, label=col, zorder=5)
            plotted += 1

        ax.axhline(100, color=HX_GRIS, linewidth=0.7, linestyle="dotted")
        ax.set_ylabel("Base 100", fontsize=7.5, color=HX_MARINE)
        ax.tick_params(colors=HX_MARINE, labelsize=7)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_color(HX_GRIS)
        ax.spines["bottom"].set_color(HX_GRIS)
        ax.grid(axis="y", alpha=0.18, color=HX_GRIS, linewidth=0.35)
        ax.grid(axis="x", visible=False)
        if plotted > 0:
            ax.legend(fontsize=7, frameon=True, framealpha=0.90,
                      edgecolor=HX_GRIS, labelcolor=HX_MARINE,
                      loc="upper left", ncol=min(3, plotted))
        ax.set_title("Evolution NAV - Base 100",
                     fontsize=9, fontweight="bold", color=HX_MARINE, pad=8)
        plt.xticks(rotation=12, ha="right", fontsize=7)

    # subplots_adjust instead of tight_layout to preserve declared figsize
    fig.subplots_adjust(left=0.08, right=0.97, top=0.90, bottom=0.14)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, facecolor=HX_BLANC)
    plt.close(fig)
    buf.seek(0)
    return buf


# ---------------------------------------------------------------------------
# STYLES REPORTLAB — definis un par un, zero dict spread
# ---------------------------------------------------------------------------

def _build_styles():
    # Couverture
    s_cover    = ParagraphStyle("cover", fontName="Helvetica", fontSize=9,
                                textColor=COL_BLANC, leading=17)
    # Titres sections — Ciel (dominant) au lieu d'orange
    s_section  = ParagraphStyle("section", fontName="Helvetica-Bold", fontSize=11,
                                textColor=COL_CIEL, spaceBefore=10, spaceAfter=4)
    s_subsect  = ParagraphStyle("subsect", fontName="Helvetica-Bold", fontSize=9,
                                textColor=COL_MARINE, spaceBefore=6, spaceAfter=3)
    # Corps
    s_body     = ParagraphStyle("body", fontName="Helvetica", fontSize=8.5,
                                textColor=COL_TEXTE, spaceAfter=4, leading=13)
    # En-tetes tableaux
    s_th       = ParagraphStyle("th",   fontName="Helvetica-Bold", fontSize=7.5,
                                textColor=COL_BLANC, alignment=TA_LEFT)
    s_th_r     = ParagraphStyle("th_r", fontName="Helvetica-Bold", fontSize=7.5,
                                textColor=COL_BLANC, alignment=TA_RIGHT)
    # Cellules
    s_td       = ParagraphStyle("td",   fontName="Helvetica", fontSize=7.5,
                                textColor=COL_MARINE, alignment=TA_LEFT)
    s_td_r     = ParagraphStyle("td_r", fontName="Helvetica", fontSize=7.5,
                                textColor=COL_MARINE, alignment=TA_RIGHT)
    s_td_g     = ParagraphStyle("td_g", fontName="Helvetica", fontSize=7.5,
                                textColor=COL_TEXTE, alignment=TA_LEFT)
    # KPI page de garde
    s_kpi_lbl  = ParagraphStyle("kpi_lbl", fontName="Helvetica", fontSize=7,
                                textColor=COL_BLANC, alignment=TA_CENTER)
    s_kpi_val  = ParagraphStyle("kpi_val", fontName="Helvetica-Bold", fontSize=15,
                                textColor=COL_BLANC, alignment=TA_CENTER)
    # Disclaimer
    s_disc     = ParagraphStyle("disc", fontName="Helvetica-Oblique", fontSize=6.5,
                                textColor=COL_TEXTE, alignment=TA_CENTER, leading=10)
    # Perf tableau
    s_pth_l    = ParagraphStyle("pth_l", fontName="Helvetica-Bold", fontSize=7.5,
                                textColor=COL_BLANC, alignment=TA_LEFT)
    s_pth_r    = ParagraphStyle("pth_r", fontName="Helvetica-Bold", fontSize=7.5,
                                textColor=COL_BLANC, alignment=TA_RIGHT)
    s_ptd_l    = ParagraphStyle("ptd_l", fontName="Helvetica", fontSize=7.5,
                                textColor=COL_MARINE, alignment=TA_LEFT)
    s_ptd      = ParagraphStyle("ptd",   fontName="Helvetica", fontSize=7.5,
                                textColor=COL_MARINE, alignment=TA_RIGHT)
    s_ppos     = ParagraphStyle("ppos",  fontName="Helvetica-Bold", fontSize=7.5,
                                textColor=COL_VERT, alignment=TA_RIGHT)
    s_pneg     = ParagraphStyle("pneg",  fontName="Helvetica-Bold", fontSize=7.5,
                                textColor=COL_ROUGE, alignment=TA_RIGHT)
    # Legende graphique
    s_caption  = ParagraphStyle("caption", fontName="Helvetica-Oblique", fontSize=7,
                                textColor=COL_TEXTE, alignment=TA_CENTER, spaceAfter=3)
    # Alerte retard (orange — seul usage dans le PDF)
    s_alert    = ParagraphStyle("alert", fontName="Helvetica-Bold", fontSize=7.5,
                                textColor=COL_ORANGE, alignment=TA_LEFT)

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
# PAGE DE GARDE — fond Marine, texte Blanc, filet Ciel
# ---------------------------------------------------------------------------

def _page_garde(styles, mode_comex, kpis, fonds_perimetre):
    elements = []
    today_str = date.today().strftime("%d %B %Y")
    perim_str = ", ".join(fonds_perimetre) if fonds_perimetre else "Tous les fonds"

    cover_text = (
        "<font color='#019ee1' size='7'>CONFIDENTIEL"
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
        ("TOPPADDING",    (0, 0), (-1, -1), 44),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 52),
        ("LEFTPADDING",   (0, 0), (-1, -1), 30),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 20),
    ]))
    elements.append(cover_tbl)
    elements.append(Spacer(1, 2))
    # Filet Ciel — couleur dominante
    elements.append(ColorRect(USABLE_W, 3.5, COL_CIEL))
    elements.append(Spacer(1, 16))

    # KPI cards — fond marine, texte blanc, separateurs Ciel
    kpi_items = [
        ("AUM Finance Total", fmt_aum(kpis.get("total_funded", 0))),
        ("Pipeline Actif",    fmt_aum(kpis.get("pipeline_actif", 0))),
        ("Taux Conversion",   "{:.1f}%".format(kpis.get("taux_conversion", 0))),
        ("Deals Actifs",      str(kpis.get("nb_deals_actifs", 0))),
    ]
    col_w = USABLE_W / 4.0
    kpi_tbl = Table(
        [[Paragraph(item[0], styles["kpi_lbl"]) for item in kpi_items],
         [Paragraph(item[1], styles["kpi_val"]) for item in kpi_items]],
        colWidths=[col_w] * 4,
        rowHeights=[0.85 * cm, 1.10 * cm],
    )
    kpi_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), COL_MARINE),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",    (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING",   (0, 0), (-1, -1), 4),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 4),
        ("LINEAFTER",     (0, 0), (2, -1),  0.4, COL_CIEL),
    ]))
    elements.append(kpi_tbl)
    elements.append(Spacer(1, 10))

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
# SECTION GRAPHIQUES
# Layout : donuts COTE A COTE (1 ligne, 2 colonnes), Top 10 dessous centré
# Tout tient sur une seule page A4 sans debordement
# ---------------------------------------------------------------------------

def _section_charts(kpis, aum_by_region, top_deals_safe, styles, mode_comex):
    """
    Layout : donuts cote a cote (1 ligne) + Top 10 dessous.
    Chaque bloc est dans un KeepTogether pour eviter tout split de page.
    Un PageBreak force garantit que les graphiques commencent sur page 2.
    """
    elements = []

    # PageBreak explicite : garantit que les charts sont sur une page propre
    elements.append(PageBreak())
    elements.append(Paragraph("Analyse du Pipeline", styles["section"]))
    elements.append(ColorRect(USABLE_W, 2, COL_CIEL))
    elements.append(Spacer(1, 12))

    # --- Bloc 1 : 2 donuts cote a cote — dans un KeepTogether ---
    d_w_in = DONUT_W / 72.0
    d_h_in = DONUT_H / 72.0

    abt    = kpis.get("aum_by_type", {})
    buf_d1 = _make_donut_png(
        list(abt.keys()), list(abt.values()),
        "AUM par Type de Client", d_w_in, d_h_in
    )
    img_d1 = Image(buf_d1, width=DONUT_W, height=DONUT_H)

    buf_d2 = _make_donut_png(
        list(aum_by_region.keys()), list(aum_by_region.values()),
        "AUM par Region Geographique", d_w_in, d_h_in
    )
    img_d2 = Image(buf_d2, width=DONUT_W, height=DONUT_H)

    gap = USABLE_W - 2 * DONUT_W
    tbl_donuts = Table(
        [[img_d1, Spacer(gap, 1), img_d2]],
        colWidths=[DONUT_W, gap, DONUT_W],
    )
    tbl_donuts.setStyle(TableStyle([
        ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING",   (0, 0), (-1, -1), 0),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 0),
        ("TOPPADDING",    (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    cap_donuts = Paragraph(
        "Figure 1 — Repartition AUM Finance par Type de Client et par Region",
        styles["caption"]
    )
    # KeepTogether : donuts + caption => jamais coupes a cheval sur 2 pages
    elements.append(KeepTogether([tbl_donuts, cap_donuts]))
    elements.append(Spacer(1, 18))

    # --- Bloc 2 : Top 10 Deals centre — dans un KeepTogether ---
    n_deals   = len(top_deals_safe[:10])
    # Hauteur lineaire mais plafonnee pour tenir sur la page
    t10_h_pts = max(80, min(210, n_deals * 20 + 25))
    t10_w_in  = TOP10_W / 72.0
    t10_h_in  = t10_h_pts / 72.0

    buf_top = _make_top10_png(top_deals_safe, mode_comex, t10_w_in, t10_h_in)
    img_top = Image(buf_top, width=TOP10_W, height=t10_h_pts)

    tbl_top = Table(
        [[Spacer(PAD_TOP10, 1), img_top, Spacer(PAD_TOP10, 1)]],
        colWidths=[PAD_TOP10, TOP10_W, PAD_TOP10],
    )
    tbl_top.setStyle(TableStyle([
        ("LEFTPADDING",   (0, 0), (-1, -1), 0),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 0),
        ("TOPPADDING",    (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    cap_top = Paragraph("Figure 2 — Top 10 Deals par AUM Finance", styles["caption"])
    elements.append(KeepTogether([tbl_top, cap_top]))
    elements.append(Spacer(1, 8))
    return elements


# ---------------------------------------------------------------------------
# SECTION PIPELINE ACTIF
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
                overdue   = nad < today
                nad_str   = "[!] {}".format(nad.isoformat()) if overdue else nad.isoformat()
                nad_style = styles["alert"] if overdue else styles["td"]
            else:
                nad_str   = "-"
                nad_style = styles["td"]

            rows.append([
                Paragraph(str(row.get("nom_client", ""))[:26],   styles["td"]),
                Paragraph(str(row.get("fonds", "")),              styles["td"]),
                Paragraph(str(row.get("statut", "")),             styles["td"]),
                Paragraph(fmt_aum(float(row.get("target_aum_initial", 0) or 0)), styles["td_r"]),
                Paragraph(fmt_aum(float(row.get("revised_aum", 0) or 0)),         styles["td_r"]),
                Paragraph(nad_str, nad_style),
                Paragraph(str(row.get("sales_owner", ""))[:14],  styles["td"]),
            ])

        tbl = Table(rows, colWidths=col_widths, repeatRows=1)
        tbl.setStyle(TableStyle([
            ("BACKGROUND",     (0, 0), (-1, 0),  COL_MARINE),
            ("LINEBELOW",      (0, 0), (-1, 0),  1.5, COL_CIEL),
            ("GRID",           (0, 0), (-1, -1), 0.3, COL_GRIS),
            ("VALIGN",         (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING",     (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING",  (0, 0), (-1, -1), 4),
            ("LEFTPADDING",    (0, 0), (-1, -1), 5),
            ("RIGHTPADDING",   (0, 0), (-1, -1), 5),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [COL_HEADER, COL_BLANC]),
        ]))
        elements.append(tbl)

    # Tableau Lost / Paused
    lost_paused = pipeline_df[pipeline_df["statut"].isin(["Lost", "Paused"])].copy()
    if not lost_paused.empty:
        elements.append(Spacer(1, 16))
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
        lp_headers = ["Client", "Fonds", "Statut", "Raison", "Concurrent"]
        lp_rows    = [[Paragraph(h, styles["th"]) for h in lp_headers]]

        for _, row in lost_paused.iterrows():
            lp_rows.append([
                Paragraph(str(row.get("nom_client", ""))[:24],          styles["td_g"]),
                Paragraph(str(row.get("fonds", "")),                     styles["td_g"]),
                Paragraph(str(row.get("statut", "")),                    styles["td_g"]),
                Paragraph(str(row.get("raison_perte", "") or "-"),       styles["td_g"]),
                Paragraph(str(row.get("concurrent_choisi", "") or "-"),  styles["td_g"]),
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
# SECTION PERFORMANCE (optionnelle)
# ---------------------------------------------------------------------------

def _section_performance(perf_data, nav_base100_df, styles, fonds_perimetre):
    elements = []
    elements.append(PageBreak())
    elements.append(Paragraph("Analyse de Performance - NAV", styles["section"]))
    elements.append(ColorRect(USABLE_W, 2, COL_CIEL))
    elements.append(Spacer(1, 12))

    pf = perf_data.copy() if perf_data is not None else pd.DataFrame()
    nb = nav_base100_df.copy() if nav_base100_df is not None else None

    if fonds_perimetre:
        if not pf.empty and "Fonds" in pf.columns:
            pf = pf[pf["Fonds"].isin(fonds_perimetre)]
        if nb is not None and hasattr(nb, "columns"):
            cols_k = [c for c in nb.columns if c in fonds_perimetre]
            nb = nb[cols_k] if cols_k else pd.DataFrame()

    # Graphique NAV Base 100
    nav_w_in = NAV_W / 72.0
    nav_h_in = NAV_H / 72.0
    buf_nav  = _make_nav_png(nb, nav_w_in, nav_h_in)
    img_nav  = Image(buf_nav, width=NAV_W, height=NAV_H)

    row_nav  = Table([[img_nav]], colWidths=[NAV_W])
    row_nav.setStyle(TableStyle([
        ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
        ("LEFTPADDING",   (0, 0), (-1, -1), 0),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 0),
        ("TOPPADDING",    (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    cap_nav = Paragraph("Figure 3 - Evolution NAV Base 100", styles["caption"])
    elements.append(KeepTogether([row_nav, cap_nav]))
    elements.append(Spacer(1, 16))

    if pf.empty:
        elements.append(Paragraph(
            "Aucune donnee de performance disponible.", styles["body"]))
        return elements

    # Tableau performances
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

    header_row = [
        Paragraph(c, styles["pth_l"] if i == 0 else styles["pth_r"])
        for i, c in enumerate(available)
    ]
    rows = [header_row]

    for _, row in pf.iterrows():
        data_row = []
        for i, col in enumerate(available):
            val = row.get(col)
            if col == "Fonds":
                data_row.append(Paragraph(str(val), styles["ptd_l"]))
            elif col in perf_cols:
                try:
                    fval  = float(val)
                    valid = not np.isnan(fval)
                except (TypeError, ValueError):
                    valid = False
                    fval  = 0.0
                if not valid:
                    data_row.append(Paragraph("n.d.", styles["ptd"]))
                else:
                    sign = "+" if fval > 0 else ""
                    txt  = "{}{:.2f}%".format(sign, fval)
                    sty  = styles["ppos"] if fval >= 0 else styles["pneg"]
                    data_row.append(Paragraph(txt, sty))
            else:
                try:
                    fval = float(val)
                    txt  = "{:.4f}".format(fval) if "NAV" in col else "{:.2f}".format(fval)
                except (TypeError, ValueError):
                    txt = "-"
                data_row.append(Paragraph(txt, styles["ptd"]))
        rows.append(data_row)

    perf_tbl = Table(rows, colWidths=col_widths, repeatRows=1)
    perf_tbl.setStyle(TableStyle([
        ("BACKGROUND",     (0, 0), (-1, 0),  COL_MARINE),
        ("LINEBELOW",      (0, 0), (-1, 0),  1.2, COL_CIEL),
        ("GRID",           (0, 0), (-1, -1), 0.3, COL_GRIS),
        ("VALIGN",         (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",     (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING",  (0, 0), (-1, -1), 5),
        ("LEFTPADDING",    (0, 0), (-1, -1), 6),
        ("RIGHTPADDING",   (0, 0), (-1, -1), 6),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [COL_HEADER, COL_BLANC]),
    ]))
    elements.append(perf_tbl)
    return elements


# ---------------------------------------------------------------------------
# EN-TETE / PIED DE PAGE
# ---------------------------------------------------------------------------

def _header_footer(canvas, doc):
    canvas.saveState()
    if doc.page > 1:
        # Pied de page marine
        canvas.setFillColor(COL_MARINE)
        canvas.rect(0, 0, PAGE_W, 0.90 * cm, fill=1, stroke=0)
        canvas.setFont("Helvetica", 6.5)
        canvas.setFillColor(HexColor("#7ab8d8"))
        canvas.drawString(2 * cm, 0.30 * cm,
                          "Executive Report - Asset Management Division")
        canvas.drawRightString(PAGE_W - 2 * cm, 0.30 * cm,
                               "Page {}".format(doc.page))
        # Filet Ciel en haut de page
        canvas.setStrokeColor(COL_CIEL)
        canvas.setLineWidth(1.8)
        canvas.line(0, PAGE_H - 0.38 * cm, PAGE_W, PAGE_H - 0.38 * cm)
        canvas.setFont("Helvetica", 6.5)
        canvas.setFillColor(COL_MARINE)
        canvas.drawRightString(PAGE_W - 2 * cm, PAGE_H - 0.26 * cm, "CONFIDENTIEL")
    canvas.restoreState()


# ---------------------------------------------------------------------------
# ENTREE PRINCIPALE
# ---------------------------------------------------------------------------

def generate_pdf(pipeline_df, kpis, aum_by_region=None, mode_comex=False,
                 perf_data=None, nav_base100_df=None, fonds_perimetre=None):
    """
    Genere le PDF Executive Pitchbook.
    Layout graphiques : donuts cote a cote + top10 dessous — tout sur 1 page.
    Retourne bytes du PDF.
    """
    aum_by_region = aum_by_region or {}

    if mode_comex:
        pipeline_df = pipeline_df.copy()
        pipeline_df["nom_client"] = pipeline_df.apply(
            lambda r: "{} - {}".format(r.get("type_client", ""), r.get("region", "")),
            axis=1
        )
        top_deals_safe = [
            dict(d, nom_client="{} - {}".format(
                d.get("type_client", ""), d.get("region", "")))
            for d in kpis.get("top_deals", [])
        ]
        kpis = dict(kpis, top_deals=top_deals_safe)
    else:
        top_deals_safe = kpis.get("top_deals", [])

    styles  = _build_styles()
    pdf_buf = io.BytesIO()

    doc = SimpleDocTemplate(
        pdf_buf,
        pagesize=A4,
        leftMargin=MARGIN_H, rightMargin=MARGIN_H,
        topMargin=MARGIN_V,  bottomMargin=1.4 * cm,
        title="Executive Report - Asset Management",
        author="Asset Management Division",
    )

    elements = []
    elements += _page_garde(styles, mode_comex, kpis, fonds_perimetre)
    elements += _section_charts(kpis, aum_by_region, top_deals_safe, styles, mode_comex)
    elements += _section_pipeline(pipeline_df, styles, mode_comex)

    if perf_data is not None and hasattr(perf_data, "empty") and not perf_data.empty:
        elements += _section_performance(perf_data, nav_base100_df, styles, fonds_perimetre)

    doc.build(elements, onFirstPage=_header_footer, onLaterPages=_header_footer)
    result = pdf_buf.getvalue()
    pdf_buf.close()
    return result
