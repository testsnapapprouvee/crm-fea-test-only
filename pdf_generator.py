# =============================================================================
# pdf_generator.py  —  CRM Asset Management  —  Amundi Edition
# Priorite : STABILITE ABSOLUE — code plat, zero helper dict de style
# Charte : #001c4b Marine | #019ee1 Ciel | #f07d00 Orange (titres sections)
# Graphiques : CHART_W = USABLE_W * 0.65 — centrage avec padding lateral
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
    Image, PageBreak,
)
from reportlab.platypus.flowables import Flowable


# ---------------------------------------------------------------------------
# COULEURS — charte Amundi (variables directement utilisees, pas de dict)
# ---------------------------------------------------------------------------

COL_MARINE  = HexColor("#001c4b")
COL_CIEL    = HexColor("#019ee1")
COL_ORANGE  = HexColor("#f07d00")
COL_BLANC   = HexColor("#ffffff")
COL_GRIS    = HexColor("#e8e8e8")
COL_TEXTE   = HexColor("#444444")
COL_VERT    = HexColor("#1a7a3c")
COL_ROUGE   = HexColor("#8b2020")

HX_MARINE   = "#001c4b"
HX_CIEL     = "#019ee1"
HX_ORANGE   = "#f07d00"
HX_BLANC    = "#ffffff"
HX_GRIS     = "#e8e8e8"
HX_TEXTE    = "#444444"

# ---------------------------------------------------------------------------
# MISE EN PAGE
# ---------------------------------------------------------------------------

PAGE_W, PAGE_H = A4
MARGIN_H  = 2.0 * cm
MARGIN_V  = 1.8 * cm
USABLE_W  = PAGE_W - 2 * MARGIN_H    # ~481 pts

# 65 % de la largeur utile — rendu factsheet Amundi
CHART_W   = USABLE_W * 0.65

# Palette graphiques Matplotlib
PALETTE = [
    "#1a5e8a", "#001c4b", "#4a8fbd", "#003f7a",
    "#2c7fb8", "#004f8c", "#6baed6", "#08519c",
    "#9ecae1", "#003060",
]

# rcParams Matplotlib — plats, sans helper
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
    "grid.linewidth":    0.45,
}


# ---------------------------------------------------------------------------
# HELPERS SIMPLES
# ---------------------------------------------------------------------------

class ColorRect(Flowable):
    """Rectangle colore pour filets de section."""
    def __init__(self, width, height, color):
        super().__init__()
        self.width  = width
        self.height = height
        self.color  = color

    def draw(self):
        self.canv.setFillColor(self.color)
        self.canv.rect(0, 0, self.width, self.height, fill=1, stroke=0)


def fmt_aum(v):
    """Formatage M EUR / Md EUR avec 1 decimale."""
    try:
        fv = float(v)
    except (TypeError, ValueError):
        return "—"
    if fv >= 1_000_000_000:
        return "{:.1f} Md EUR".format(fv / 1_000_000_000)
    if fv >= 1_000_000:
        return "{:.1f} M EUR".format(fv / 1_000_000)
    if fv >= 1_000:
        return "{:.0f} k EUR".format(fv / 1_000)
    return "{:.0f} EUR".format(fv)


def _buf_to_centered_image(buf, aspect_ratio, elements, caption, styles):
    """
    Insere un graphique centre dans elements.
    CHART_W = USABLE_W * 0.65 — padding lateral pour centrage.
    """
    w   = CHART_W
    h   = w * aspect_ratio
    img = Image(buf, width=w, height=h)
    pad = (USABLE_W - w) / 2.0

    row_data  = [[Spacer(pad, 1), img, Spacer(pad, 1)]]
    col_widths = [pad, w, pad]
    wrapper = Table(row_data, colWidths=col_widths)
    wrapper.setStyle(TableStyle([
        ("LEFTPADDING",   (0, 0), (-1, -1), 0),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 0),
        ("TOPPADDING",    (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    elements.append(wrapper)
    elements.append(Paragraph(caption, styles["caption"]))
    elements.append(Spacer(1, 8))
    elements.append(ColorRect(USABLE_W * 0.35, 1, COL_GRIS))
    elements.append(Spacer(1, 14))


# ---------------------------------------------------------------------------
# GRAPHIQUES MATPLOTLIB — chaque fonction est autonome et appelle plt.close()
# ---------------------------------------------------------------------------

def _make_donut(labels, values, title):
    """Donut chart Amundi — retourne io.BytesIO PNG."""
    plt.rcParams.update(MPL_RC)
    fig_w_in = CHART_W / 72.0
    fig, ax = plt.subplots(figsize=(fig_w_in, fig_w_in * 0.82))
    fig.patch.set_facecolor(HX_BLANC)

    if not labels or sum(values) == 0:
        ax.text(0.5, 0.5, "Aucune donnee", ha="center", va="center",
                fontsize=9, color=HX_MARINE, transform=ax.transAxes)
        ax.axis("off")
    else:
        colors = [PALETTE[i % len(PALETTE)] for i in range(len(labels))]
        _, _, autotexts = ax.pie(
            values, colors=colors, autopct="%1.1f%%", startangle=90,
            pctdistance=0.74,
            wedgeprops={"width": 0.52, "edgecolor": HX_BLANC, "linewidth": 1.8}
        )
        for at in autotexts:
            at.set_fontsize(8.5)
            at.set_color(HX_BLANC)
            at.set_fontweight("bold")

        total = sum(values)
        ax.text(0,  0.09, fmt_aum(total), ha="center", va="center",
                fontsize=9, fontweight="bold", color=HX_MARINE)
        ax.text(0, -0.19, "AUM Finance", ha="center", va="center",
                fontsize=7, color=HX_TEXTE)

        patches = [mpatches.Patch(color=colors[i],
                   label="{}: {}".format(labels[i], fmt_aum(values[i])))
                   for i in range(len(labels))]
        ax.legend(handles=patches, loc="lower center",
                  bbox_to_anchor=(0.5, -0.30), ncol=2,
                  fontsize=7, frameon=False, labelcolor=HX_MARINE)
        ax.set_title(title, fontsize=9.5, fontweight="bold",
                     color=HX_MARINE, pad=10)

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight", facecolor=HX_BLANC)
    plt.close(fig)
    buf.seek(0)
    return buf


def _make_top10_bar(top_deals, mode_comex):
    """Bar chart horizontal Top 10 deals — retourne io.BytesIO PNG."""
    plt.rcParams.update(MPL_RC)
    deals  = top_deals[:10]
    n      = len(deals)
    fig_w  = CHART_W / 72.0
    fig_h  = max(2.8, n * 0.52 + 0.9)
    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    fig.patch.set_facecolor(HX_BLANC)
    ax.set_facecolor(HX_BLANC)

    if not deals:
        ax.text(0.5, 0.5, "Aucun deal Funded", ha="center", va="center",
                fontsize=9, color=HX_MARINE, transform=ax.transAxes)
        ax.axis("off")
    else:
        labels = [
            "{} - {}".format(d["type_client"], d["region"]) if mode_comex
            else str(d.get("nom_client", ""))
            for d in deals
        ]
        values = [float(d.get("funded_aum", 0)) for d in deals]
        pairs  = sorted(zip(values, labels), reverse=True)
        values = [v for v, _ in pairs]
        labels = [l for _, l in pairs]
        max_v  = max(values) if values else 1.0
        colors = [PALETTE[i % len(PALETTE)] for i in range(n)]

        bars = ax.barh(range(len(labels)), values, color=colors,
                       edgecolor=HX_BLANC, height=0.62, linewidth=0.5)
        for bar, val in zip(bars, values):
            ax.text(bar.get_width() + max_v * 0.012,
                    bar.get_y() + bar.get_height() / 2,
                    fmt_aum(val), va="center", ha="left",
                    fontsize=7.5, color=HX_MARINE, fontweight="bold")

        ax.set_yticks(range(len(labels)))
        ax.set_yticklabels([l[:24] for l in labels], fontsize=8, color=HX_MARINE)
        ax.invert_yaxis()
        ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: fmt_aum(x)))
        ax.tick_params(axis="x", labelsize=7, colors=HX_MARINE)
        ax.set_xlim(0, max_v * 1.28)
        ax.grid(axis="x", alpha=0.22, color=HX_GRIS, linewidth=0.4)
        ax.spines["left"].set_color(HX_GRIS)
        ax.spines["bottom"].set_color(HX_GRIS)
        ax.set_title("Top 10 Deals - AUM Finance", fontsize=9.5,
                     fontweight="bold", color=HX_MARINE, pad=9)

    fig.tight_layout(pad=0.8)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight", facecolor=HX_BLANC)
    plt.close(fig)
    buf.seek(0)
    return buf


def _make_nav_base100(nav_df):
    """Courbe NAV Base 100 — retourne io.BytesIO PNG."""
    plt.rcParams.update(MPL_RC)
    fig_w = CHART_W / 72.0
    fig_h = fig_w * 0.48
    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    fig.patch.set_facecolor(HX_BLANC)
    ax.set_facecolor(HX_BLANC)

    if nav_df is None or not hasattr(nav_df, "columns") or nav_df.empty:
        ax.text(0.5, 0.5, "Donnees NAV non disponibles",
                ha="center", va="center", fontsize=9, color=HX_MARINE,
                transform=ax.transAxes)
        ax.axis("off")
    else:
        dash_styles = ["solid", "dashed", "dotted", "dashdot", "solid"]
        plotted = 0
        for i, col in enumerate(nav_df.columns):
            series = nav_df[col].dropna()
            if series.empty:
                continue
            color = PALETTE[i % len(PALETTE)]
            dash  = dash_styles[i % len(dash_styles)]
            if len(series) >= 2:
                ax.plot(series.index, series.values, label=col,
                        color=color, linewidth=1.6, linestyle=dash, alpha=0.92)
            else:
                ax.scatter(series.index, series.values, color=color,
                           s=55, label=col, zorder=5)
            ax.scatter([series.index[-1]], [series.values[-1]],
                       color=color, s=28, zorder=6)
            plotted += 1

        ax.axhline(100, color=HX_GRIS, linewidth=0.7, linestyle="dotted")
        ax.set_ylabel("Base 100", fontsize=8, color=HX_MARINE)
        ax.tick_params(colors=HX_MARINE, labelsize=7.5)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_color(HX_GRIS)
        ax.spines["bottom"].set_color(HX_GRIS)
        ax.grid(axis="y", alpha=0.20, color=HX_GRIS, linewidth=0.4)
        ax.grid(axis="x", visible=False)
        if plotted > 0:
            ax.legend(fontsize=7.5, frameon=True, framealpha=0.92,
                      edgecolor=HX_GRIS, labelcolor=HX_MARINE,
                      loc="upper left", ncol=min(3, plotted))
        ax.set_title("Evolution NAV - Base 100", fontsize=9.5,
                     fontweight="bold", color=HX_MARINE, pad=9)
        plt.xticks(rotation=15, ha="right", fontsize=7)

    fig.tight_layout(pad=0.7)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight", facecolor=HX_BLANC)
    plt.close(fig)
    buf.seek(0)
    return buf


# ---------------------------------------------------------------------------
# STYLES REPORTLAB — definis un par un, zero dict helper
# ---------------------------------------------------------------------------

def _build_styles():
    # Paragraphe de base
    style_cover = ParagraphStyle(
        "cover", fontName="Helvetica", fontSize=9,
        textColor=COL_BLANC, leading=17
    )
    # Titre de section — Orange Amundi
    style_section = ParagraphStyle(
        "section", fontName="Helvetica-Bold", fontSize=11.5,
        textColor=COL_ORANGE, spaceBefore=12, spaceAfter=5
    )
    # Corps de texte
    style_body = ParagraphStyle(
        "body", fontName="Helvetica", fontSize=8.5,
        textColor=COL_TEXTE, spaceAfter=4, leading=13
    )
    # En-tetes tableaux (fond marine = couleur apliquee via TableStyle)
    style_th = ParagraphStyle(
        "th", fontName="Helvetica-Bold", fontSize=7.5,
        textColor=COL_BLANC, alignment=TA_LEFT
    )
    style_th_r = ParagraphStyle(
        "th_r", fontName="Helvetica-Bold", fontSize=7.5,
        textColor=COL_BLANC, alignment=TA_RIGHT
    )
    # Cellules tableau
    style_td = ParagraphStyle(
        "td", fontName="Helvetica", fontSize=7.5,
        textColor=COL_MARINE, alignment=TA_LEFT
    )
    style_td_r = ParagraphStyle(
        "td_r", fontName="Helvetica", fontSize=7.5,
        textColor=COL_MARINE, alignment=TA_RIGHT
    )
    style_td_g = ParagraphStyle(
        "td_g", fontName="Helvetica", fontSize=7.5,
        textColor=COL_TEXTE, alignment=TA_LEFT
    )
    # KPI page de garde
    style_kpi_lbl = ParagraphStyle(
        "kpi_lbl", fontName="Helvetica", fontSize=7,
        textColor=COL_BLANC, alignment=TA_CENTER
    )
    style_kpi_val = ParagraphStyle(
        "kpi_val", fontName="Helvetica-Bold", fontSize=15,
        textColor=COL_BLANC, alignment=TA_CENTER
    )
    # Disclaimer
    style_disc = ParagraphStyle(
        "disc", fontName="Helvetica-Oblique", fontSize=6.5,
        textColor=COL_TEXTE, alignment=TA_CENTER, leading=10
    )
    # Tableau performances
    style_perf_th_l = ParagraphStyle(
        "perf_th_l", fontName="Helvetica-Bold", fontSize=7.5,
        textColor=COL_BLANC, alignment=TA_LEFT
    )
    style_perf_th_r = ParagraphStyle(
        "perf_th_r", fontName="Helvetica-Bold", fontSize=7.5,
        textColor=COL_BLANC, alignment=TA_RIGHT
    )
    style_perf_td_l = ParagraphStyle(
        "perf_td_l", fontName="Helvetica", fontSize=7.5,
        textColor=COL_MARINE, alignment=TA_LEFT
    )
    style_perf_td = ParagraphStyle(
        "perf_td", fontName="Helvetica", fontSize=7.5,
        textColor=COL_MARINE, alignment=TA_RIGHT
    )
    style_perf_pos = ParagraphStyle(
        "perf_pos", fontName="Helvetica-Bold", fontSize=7.5,
        textColor=COL_VERT, alignment=TA_RIGHT
    )
    style_perf_neg = ParagraphStyle(
        "perf_neg", fontName="Helvetica-Bold", fontSize=7.5,
        textColor=COL_ROUGE, alignment=TA_RIGHT
    )
    # Caption graphique
    style_caption = ParagraphStyle(
        "caption", fontName="Helvetica-Oblique", fontSize=7,
        textColor=COL_TEXTE, alignment=TA_CENTER, spaceAfter=4
    )
    # Alerte retard dans tableau pipeline
    style_alert = ParagraphStyle(
        "alert", fontName="Helvetica-Bold", fontSize=7.5,
        textColor=COL_ORANGE, alignment=TA_LEFT
    )

    return {
        "cover":       style_cover,
        "section":     style_section,
        "body":        style_body,
        "th":          style_th,
        "th_r":        style_th_r,
        "td":          style_td,
        "td_r":        style_td_r,
        "td_g":        style_td_g,
        "kpi_lbl":     style_kpi_lbl,
        "kpi_val":     style_kpi_val,
        "disc":        style_disc,
        "perf_th_l":   style_perf_th_l,
        "perf_th_r":   style_perf_th_r,
        "perf_td_l":   style_perf_td_l,
        "perf_td":     style_perf_td,
        "perf_pos":    style_perf_pos,
        "perf_neg":    style_perf_neg,
        "caption":     style_caption,
        "alert":       style_alert,
    }


# ---------------------------------------------------------------------------
# PAGE DE GARDE — fond Marine, texte Blanc
# ---------------------------------------------------------------------------

def _page_garde(styles, mode_comex, kpis, fonds_perimetre):
    elements = []
    today_str = date.today().strftime("%d %B %Y")
    perim_str = ", ".join(fonds_perimetre) if fonds_perimetre else "Tous les fonds"

    # Bloc couverture
    cover_text = (
        "<font color='#f07d00' size='7'>CONFIDENTIEL"
        + ("  |  MODE COMEX ACTIF" if mode_comex else "") +
        "</font><br/><br/>"
        "<font color='white' size='22'><b>Executive Report</b></font><br/>"
        "<font color='white' size='15'>Pipeline &amp; Reporting</font><br/><br/>"
        "<font color='#b0c8dc' size='8.5'>Asset Management Division</font><br/>"
        "<font color='white' size='7.5'>Perimetre : " + perim_str + "</font><br/><br/>"
        "<font color='#b0c8dc' size='7.5'>" + today_str + "</font>"
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
    elements.append(ColorRect(USABLE_W, 3.5, COL_ORANGE))
    elements.append(Spacer(1, 16))

    # KPI cards
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
        rowHeights=[0.9 * cm, 1.15 * cm]
    )
    kpi_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), COL_MARINE),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",    (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ("LEFTPADDING",   (0, 0), (-1, -1), 4),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 4),
        ("LINEAFTER",     (0, 0), (2, -1),  0.4, COL_ORANGE),
    ]))
    elements.append(kpi_tbl)
    elements.append(Spacer(1, 10))

    disc_text = (
        "Document strictement confidentiel a usage interne exclusif. "
        "Reproduction et diffusion externe interdites. "
        "Les performances passees ne prejudgent pas des performances futures."
    )
    if mode_comex:
        disc_text += " Mode Comex actif : noms clients anonymises."
    elements.append(Paragraph(disc_text, styles["disc"]))
    elements.append(PageBreak())
    return elements


# ---------------------------------------------------------------------------
# SECTION GRAPHIQUES — ordre strict : 1.Type | 2.Region | 3.Top10
# ---------------------------------------------------------------------------

def _section_charts(kpis, aum_by_region, top_deals_safe, styles, mode_comex):
    elements = []
    elements.append(Paragraph("Analyse du Pipeline", styles["section"]))
    elements.append(ColorRect(USABLE_W, 2.5, COL_ORANGE))
    elements.append(Spacer(1, 14))

    # 1 — Type de Client
    abt = kpis.get("aum_by_type", {})
    buf1 = _make_donut(list(abt.keys()), list(abt.values()),
                       "Repartition AUM Funded - Type de Client")
    _buf_to_centered_image(buf1, 0.82,
                           elements,
                           "Figure 1 - Repartition AUM Finance par Type de Client",
                           styles)

    # 2 — Region Geographique
    buf2 = _make_donut(list(aum_by_region.keys()), list(aum_by_region.values()),
                       "Repartition AUM Funded - Region Geographique")
    _buf_to_centered_image(buf2, 0.82,
                           elements,
                           "Figure 2 - Repartition AUM Finance par Region Geographique",
                           styles)

    # 3 — Top 10 Deals
    n_deals = len(top_deals_safe[:10])
    aspect  = max(0.44, n_deals * 0.060 + 0.18)
    buf3 = _make_top10_bar(top_deals_safe, mode_comex)
    _buf_to_centered_image(buf3, aspect,
                           elements,
                           "Figure 3 - Top 10 Deals par AUM Finance",
                           styles)
    return elements


# ---------------------------------------------------------------------------
# SECTION PIPELINE
# ---------------------------------------------------------------------------

def _section_pipeline(pipeline_df, styles, mode_comex):
    elements = []
    elements.append(PageBreak())
    elements.append(Paragraph("Pipeline Actif - Recapitulatif", styles["section"]))
    elements.append(ColorRect(USABLE_W, 2.5, COL_ORANGE))
    elements.append(Spacer(1, 10))

    actifs = pipeline_df[pipeline_df["statut"].isin(
        ["Prospect", "Initial Pitch", "Due Diligence", "Soft Commit"]
    )].copy()

    if actifs.empty:
        elements.append(Paragraph("Aucun deal actif dans le perimetre.", styles["body"]))
    else:
        if mode_comex:
            actifs["nom_client"] = actifs.apply(
                lambda r: "{} - {}".format(r["type_client"], r["region"]), axis=1
            )

        # Ratios colonnes — somme = 1.0
        ratios     = [0.22, 0.16, 0.12, 0.13, 0.13, 0.13, 0.11]
        col_widths = [USABLE_W * r for r in ratios]
        headers    = ["Client", "Fonds", "Statut", "AUM Cible",
                      "AUM Revise", "Prochaine Action", "Commercial"]

        rows  = [[Paragraph(h, styles["th"]) for h in headers]]
        today = date.today()

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
                Paragraph(str(row.get("nom_client", ""))[:26],  styles["td"]),
                Paragraph(str(row.get("fonds", "")),             styles["td"]),
                Paragraph(str(row.get("statut", "")),            styles["td"]),
                Paragraph(fmt_aum(float(row.get("target_aum_initial", 0) or 0)), styles["td_r"]),
                Paragraph(fmt_aum(float(row.get("revised_aum", 0) or 0)),         styles["td_r"]),
                Paragraph(nad_str, nad_style),
                Paragraph(str(row.get("sales_owner", ""))[:14],  styles["td"]),
            ])

        tbl = Table(rows, colWidths=col_widths, repeatRows=1)
        tbl.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (-1, 0),  COL_MARINE),
            ("LINEBELOW",     (0, 0), (-1, 0),  1.5, COL_ORANGE),
            ("GRID",          (0, 0), (-1, -1), 0.3, COL_GRIS),
            ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING",    (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING",   (0, 0), (-1, -1), 5),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 5),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [HexColor("#f7fafd"), COL_BLANC]),
        ]))
        elements.append(tbl)

    # Tableau Lost / Paused
    lost_paused = pipeline_df[pipeline_df["statut"].isin(["Lost", "Paused"])].copy()
    if not lost_paused.empty:
        elements.append(Spacer(1, 16))
        elements.append(Paragraph("Deals Perdus / En Pause", styles["section"]))
        elements.append(ColorRect(USABLE_W, 1.5, COL_GRIS))
        elements.append(Spacer(1, 8))

        if mode_comex:
            lost_paused["nom_client"] = lost_paused.apply(
                lambda r: "{} - {}".format(r["type_client"], r["region"]), axis=1
            )

        lp_ratios  = [0.26, 0.18, 0.12, 0.22, 0.22]
        lp_col_w   = [USABLE_W * r for r in lp_ratios]
        lp_headers = ["Client", "Fonds", "Statut", "Raison", "Concurrent"]
        lp_rows    = [[Paragraph(h, styles["th"]) for h in lp_headers]]

        for _, row in lost_paused.iterrows():
            lp_rows.append([
                Paragraph(str(row.get("nom_client", ""))[:24],            styles["td_g"]),
                Paragraph(str(row.get("fonds", "")),                      styles["td_g"]),
                Paragraph(str(row.get("statut", "")),                     styles["td_g"]),
                Paragraph(str(row.get("raison_perte", "") or "-"),        styles["td_g"]),
                Paragraph(str(row.get("concurrent_choisi", "") or "-"),   styles["td_g"]),
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
    elements.append(ColorRect(USABLE_W, 2.5, COL_ORANGE))
    elements.append(Spacer(1, 12))

    # Filtrage sur le perimetre
    pf = perf_data.copy() if perf_data is not None else pd.DataFrame()
    nb = nav_base100_df.copy() if nav_base100_df is not None else None

    if fonds_perimetre:
        if not pf.empty and "Fonds" in pf.columns:
            pf = pf[pf["Fonds"].isin(fonds_perimetre)]
        if nb is not None and hasattr(nb, "columns"):
            cols_k = [col for col in nb.columns if col in fonds_perimetre]
            nb = nb[cols_k] if cols_k else pd.DataFrame()

    # Graphique NAV Base 100
    buf_nav = _make_nav_base100(nb)
    _buf_to_centered_image(buf_nav, 0.50, elements,
                           "Figure 4 - Evolution NAV Base 100",
                           styles)

    if pf.empty:
        elements.append(Paragraph(
            "Aucune donnee de performance disponible pour ce perimetre.",
            styles["body"]
        ))
        return elements

    # Tableau performances
    elements.append(Paragraph("Tableau des Performances", styles["section"]))
    elements.append(ColorRect(USABLE_W, 2.5, COL_ORANGE))
    elements.append(Spacer(1, 10))

    perf_col_names = {"Perf 1M (%)", "Perf YTD (%)", "Perf Periode (%)"}
    col_order = ["Fonds", "NAV Derniere", "Base 100 Actuel",
                 "Perf 1M (%)", "Perf YTD (%)", "Perf Periode (%)"]
    available = [col for col in col_order if col in pf.columns]
    n = len(available)

    if n == 0:
        elements.append(Paragraph("Colonnes manquantes dans les donnees.", styles["body"]))
        return elements

    fonds_r    = 0.28
    rest_r     = (1.0 - fonds_r) / max(n - 1, 1)
    col_widths = [USABLE_W * (fonds_r if i == 0 else rest_r) for i in range(n)]

    header_row = []
    for i, col in enumerate(available):
        sty = styles["perf_th_l"] if i == 0 else styles["perf_th_r"]
        header_row.append(Paragraph(col, sty))
    rows = [header_row]

    for _, row in pf.iterrows():
        data_row = []
        for i, col in enumerate(available):
            val = row.get(col)
            if col == "Fonds":
                data_row.append(Paragraph(str(val), styles["perf_td_l"]))
            elif col in perf_col_names:
                try:
                    fval = float(val)
                    valid = not np.isnan(fval)
                except (TypeError, ValueError):
                    valid = False
                    fval  = 0.0
                if not valid:
                    data_row.append(Paragraph("n.d.", styles["perf_td"]))
                else:
                    sign = "+" if fval > 0 else ""
                    txt  = "{}{:.2f}%".format(sign, fval)
                    sty  = styles["perf_pos"] if fval >= 0 else styles["perf_neg"]
                    data_row.append(Paragraph(txt, sty))
            else:
                try:
                    fval = float(val)
                    txt  = "{:.4f}".format(fval) if "NAV" in col else "{:.2f}".format(fval)
                except (TypeError, ValueError):
                    txt = "-"
                data_row.append(Paragraph(txt, styles["perf_td"]))
        rows.append(data_row)

    perf_tbl = Table(rows, colWidths=col_widths, repeatRows=1)
    perf_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, 0),  COL_MARINE),
        ("LINEBELOW",     (0, 0), (-1, 0),  1.5, COL_ORANGE),
        ("GRID",          (0, 0), (-1, -1), 0.3, COL_GRIS),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",    (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING",   (0, 0), (-1, -1), 6),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 6),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [HexColor("#f7fafd"), COL_BLANC]),
    ]))
    elements.append(perf_tbl)
    return elements


# ---------------------------------------------------------------------------
# EN-TETE / PIED DE PAGE
# ---------------------------------------------------------------------------

def _header_footer(canvas, doc):
    canvas.saveState()
    if doc.page > 1:
        canvas.setFillColor(COL_MARINE)
        canvas.rect(0, 0, PAGE_W, 0.95 * cm, fill=1, stroke=0)
        canvas.setFont("Helvetica", 6.5)
        canvas.setFillColor(HexColor("#b0c8dc"))
        canvas.drawString(2 * cm, 0.32 * cm, "Executive Report - Asset Management Division")
        canvas.drawRightString(PAGE_W - 2 * cm, 0.32 * cm, "Page {}".format(doc.page))
        canvas.setStrokeColor(COL_ORANGE)
        canvas.setLineWidth(2.0)
        canvas.line(0, PAGE_H - 0.40 * cm, PAGE_W, PAGE_H - 0.40 * cm)
        canvas.setFont("Helvetica", 6.5)
        canvas.setFillColor(COL_MARINE)
        canvas.drawRightString(PAGE_W - 2 * cm, PAGE_H - 0.28 * cm, "CONFIDENTIEL")
    canvas.restoreState()


# ---------------------------------------------------------------------------
# ENTREE PRINCIPALE
# ---------------------------------------------------------------------------

def generate_pdf(pipeline_df, kpis, aum_by_region=None, mode_comex=False,
                 perf_data=None, nav_base100_df=None, fonds_perimetre=None):
    """
    Genere le PDF Executive Pitchbook complet.
    Filtrage 100% pilote par fonds_perimetre (depuis la sidebar de app.py).
    Retourne bytes du PDF.
    """
    aum_by_region = aum_by_region or {}

    if mode_comex:
        pipeline_df = pipeline_df.copy()
        pipeline_df["nom_client"] = pipeline_df.apply(
            lambda r: "{} - {}".format(r.get("type_client", ""), r.get("region", "")), axis=1
        )
        top_deals_safe = [
            dict(d, nom_client="{} - {}".format(d.get("type_client", ""), d.get("region", "")))
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
        topMargin=MARGIN_V,  bottomMargin=1.5 * cm,
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
