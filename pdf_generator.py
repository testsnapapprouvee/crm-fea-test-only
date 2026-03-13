# =============================================================================
# pdf_generator.py — Rapport PDF Pitchbook — FEA/Amundi Edition Strategique
# Charte : Marine #002D54 | Ciel #00A8E1 | USABLE_W = PAGE_W - 2*marges
# Missions :
#   - kpi_value couleur BLANC (fix esthetique)
#   - Graphiques empiles verticalement, USABLE_W * 0.8
#   - Ordre : Type Client → Region → Top 10 Deals
#   - Nouveau graphique _chart_pie_aum_by_region
#   - Page Performance optionnelle (perf_data + nav_base100_df)
#   - Filtrage fund-by-fund transmis depuis app.py
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
# CONSTANTES MISE EN PAGE & CHARTE
# ---------------------------------------------------------------------------

BLEU_MARINE = HexColor("#002D54")
BLEU_CIEL   = HexColor("#00A8E1")
BLANC       = HexColor("#FFFFFF")
GRIS_CLAIR  = HexColor("#E0E0E0")
GRIS_TEXTE  = HexColor("#555555")
BLEU_MID    = HexColor("#1A6B9A")

HEX_MARINE  = "#002D54"
HEX_CIEL    = "#00A8E1"
HEX_GRIS    = "#E0E0E0"
HEX_GTXT    = "#555555"
HEX_BLANC   = "#FFFFFF"

PAGE_W, PAGE_H = A4
MARGIN_H       = 2 * cm
MARGIN_V       = 1.8 * cm
USABLE_W       = PAGE_W - 2 * MARGIN_H     # 481.89 pts — largeur stricte

# Largeur des graphiques dans le PDF (empiles, 80% de USABLE_W)
CHART_W        = USABLE_W * 0.80

# Palette Amundi — variations de bleu, zéro couleur Matplotlib par defaut
PALETTE = [
    HEX_CIEL,    # Bleu ciel Amundi
    HEX_MARINE,  # Bleu marine Amundi
    "#1A6B9A",   # Bleu intermediaire
    "#7BC8E8",   # Bleu pale
    "#003F7A",   # Bleu profond
    "#5BA3C9",   # Bleu moyen-clair
    "#2C8FBF",   # Bleu ocean
    "#004F8C",   # Bleu institutionnel
    "#A8D8EE",   # Bleu tres pale
    "#003060",   # Bleu nuit
]

MPL_PARAMS = {
    "font.family":       "DejaVu Sans",
    "axes.spines.top":   False,
    "axes.spines.right": False,
    "axes.labelcolor":   HEX_MARINE,
    "xtick.color":       HEX_MARINE,
    "ytick.color":       HEX_MARINE,
    "figure.facecolor":  HEX_BLANC,
    "axes.facecolor":    HEX_BLANC,
    "grid.color":        HEX_GRIS,
    "grid.linewidth":    0.5,
}


# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------

class ColorRect(Flowable):
    """Rectangle colore pour bandeaux de section."""
    def __init__(self, width, height, color):
        super().__init__()
        self.width  = width
        self.height = height
        self.color  = color

    def draw(self):
        self.canv.setFillColor(self.color)
        self.canv.rect(0, 0, self.width, self.height, fill=1, stroke=0)


def anonymize_df(df: pd.DataFrame) -> pd.DataFrame:
    """Mode Comex : remplace nom_client par type_client - region."""
    df = df.copy()
    if all(c in df.columns for c in ("nom_client", "type_client", "region")):
        df["nom_client"] = df.apply(
            lambda r: f"{r['type_client']} - {r['region']}", axis=1
        )
    return df


def fmt_aum(value: float) -> str:
    if value >= 1_000_000_000: return f"{value/1_000_000_000:.1f}Md EUR"
    if value >= 1_000_000:     return f"{value/1_000_000:.1f}M EUR"
    if value >= 1_000:         return f"{value/1_000:.0f}k EUR"
    return f"{value:.0f} EUR"


def _donut(ax, labels, values, title: str):
    """
    Sous-routine commune pour les deux donuts (type client & region).
    Charte Amundi stricte, legendes compactes.
    """
    colors   = [PALETTE[i % len(PALETTE)] for i in range(len(labels))]
    _, _, autotexts = ax.pie(
        values, colors=colors, autopct="%1.1f%%", startangle=90,
        pctdistance=0.74,
        wedgeprops={"width": 0.54, "edgecolor": HEX_BLANC, "linewidth": 1.5}
    )
    for at in autotexts:
        at.set_fontsize(8.5); at.set_color(HEX_BLANC); at.set_fontweight("bold")

    total = sum(values)
    ax.text(0,  0.08, fmt_aum(total),
            ha="center", va="center", fontsize=9, fontweight="bold",
            color=HEX_MARINE)
    ax.text(0, -0.18, "AUM Finance",
            ha="center", va="center", fontsize=7, color=HEX_GTXT)

    patches = [mpatches.Patch(color=colors[i],
               label=f"{labels[i]}: {fmt_aum(values[i])}")
               for i in range(len(labels))]
    ax.legend(handles=patches, loc="lower center",
              bbox_to_anchor=(0.5, -0.28), ncol=2,
              fontsize=7, frameon=False, labelcolor=HEX_MARINE)
    ax.set_title(title, fontsize=9.5, fontweight="bold",
                 color=HEX_MARINE, pad=8)


# ---------------------------------------------------------------------------
# GRAPHIQUES MATPLOTLIB
# ---------------------------------------------------------------------------

def _chart_pie_aum_by_type(aum_by_type: dict) -> io.BytesIO:
    """Donut — AUM Funded par type de client."""
    plt.rcParams.update(MPL_PARAMS)
    fig, ax = plt.subplots(figsize=(CHART_W / 72, CHART_W / 72 * 0.72))

    if not aum_by_type or sum(aum_by_type.values()) == 0:
        ax.text(0.5, 0.5, "Aucune donnee Funded",
                ha="center", va="center", color=HEX_MARINE)
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
    """Donut — AUM Funded par region geographique."""
    plt.rcParams.update(MPL_PARAMS)
    fig, ax = plt.subplots(figsize=(CHART_W / 72, CHART_W / 72 * 0.72))

    if not aum_by_region or sum(aum_by_region.values()) == 0:
        ax.text(0.5, 0.5, "Aucune donnee Funded par region",
                ha="center", va="center", color=HEX_MARINE)
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
    """Bar chart horizontal — Top 10 deals funded."""
    plt.rcParams.update(MPL_PARAMS)

    n_deals = len(top_deals[:10])
    fig_h   = max(3.0, n_deals * 0.46 + 0.8)
    fig, ax = plt.subplots(figsize=(CHART_W / 72, fig_h))

    if not top_deals:
        ax.text(0.5, 0.5, "Aucun deal Funded",
                ha="center", va="center", color=HEX_MARINE)
        ax.axis("off")
    else:
        labels = [
            f"{d['type_client']} - {d['region']}" if mode_comex
            else d["nom_client"]
            for d in top_deals[:10]
        ]
        values = [float(d["funded_aum"]) for d in top_deals[:10]]
        pairs  = sorted(zip(values, labels), reverse=True)
        values = [v for v, _ in pairs]
        labels = [l for _, l in pairs]
        max_v  = max(values) if values else 1

        bars = ax.barh(range(len(labels)), values,
                       color=HEX_CIEL, edgecolor=HEX_BLANC, height=0.6)
        for bar, val in zip(bars, values):
            ax.text(bar.get_width() + max_v * 0.01,
                    bar.get_y() + bar.get_height() / 2,
                    fmt_aum(val), va="center", ha="left",
                    fontsize=7.5, color=HEX_MARINE, fontweight="bold")

        ax.set_yticks(range(len(labels)))
        ax.set_yticklabels(labels, fontsize=8, color=HEX_MARINE)
        ax.invert_yaxis()
        ax.xaxis.set_major_formatter(
            plt.FuncFormatter(lambda x, _: fmt_aum(x))
        )
        ax.tick_params(axis="x", labelsize=7.5, colors=HEX_MARINE)
        ax.set_xlim(0, max_v * 1.22)
        ax.grid(axis="x", alpha=0.28, color=HEX_GRIS)
        ax.set_title("Top 10 Deals — AUM Finance",
                     fontsize=9.5, fontweight="bold",
                     color=HEX_MARINE, pad=8)

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150,
                bbox_inches="tight", facecolor=HEX_BLANC)
    plt.close(fig)
    buf.seek(0)
    return buf


def _chart_nav_base100(nav_base100_df: pd.DataFrame) -> io.BytesIO:
    """Courbe NAV Base 100 — insertion page Performance."""
    plt.rcParams.update(MPL_PARAMS)
    fig, ax = plt.subplots(figsize=(USABLE_W / 72, USABLE_W / 72 * 0.38))

    if nav_base100_df is None or nav_base100_df.empty:
        ax.text(0.5, 0.5, "Donnees NAV non disponibles",
                ha="center", va="center", color=HEX_MARINE)
        ax.axis("off")
    else:
        for i, fonds in enumerate(nav_base100_df.columns):
            series = nav_base100_df[fonds].dropna()
            if series.empty:
                continue
            color    = PALETTE[i % len(PALETTE)]
            line_sty = "-" if i % 2 == 0 else "--"
            if len(series) >= 2:
                ax.plot(series.index, series.values,
                        label=fonds, color=color,
                        linewidth=1.8, linestyle=line_sty, alpha=0.9)
            else:
                ax.scatter(series.index, series.values,
                           color=color, s=60, label=fonds, zorder=5)
            if not series.empty:
                ax.scatter([series.index[-1]], [series.values[-1]],
                           color=color, s=30, zorder=6)

        ax.axhline(100, color=HEX_GRIS, linewidth=0.8, linestyle=":")
        ax.set_ylabel("NAV (Base 100)", fontsize=8, color=HEX_MARINE)
        ax.tick_params(colors=HEX_MARINE, labelsize=7.5)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.grid(axis="y", alpha=0.22, color=HEX_GRIS)
        ax.legend(fontsize=7.5, frameon=True, framealpha=0.9,
                  edgecolor=HEX_GRIS, labelcolor=HEX_MARINE,
                  loc="upper left", ncol=min(3, len(nav_base100_df.columns)))
        ax.set_title("Evolution NAV — Base 100",
                     fontsize=9.5, fontweight="bold", color=HEX_MARINE, pad=8)
        plt.xticks(rotation=16, ha="right", fontsize=7)

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
        "cover_inner": ps("cover_inner",
            fontName="Helvetica", fontSize=9, textColor=BLANC, leading=17),
        "section":     ps("section",
            fontName="Helvetica-Bold", fontSize=12, textColor=BLEU_MARINE,
            spaceBefore=12, spaceAfter=6),
        "body":        ps("body",
            fontName="Helvetica", fontSize=8.5, textColor=GRIS_TEXTE,
            spaceAfter=4, leading=13),
        "th":          ps("th",
            fontName="Helvetica-Bold", fontSize=7.5, textColor=BLANC),
        "td":          ps("td",
            fontName="Helvetica", fontSize=7.5, textColor=BLEU_MARINE),
        "td_grey":     ps("td_grey",
            fontName="Helvetica", fontSize=7.5, textColor=GRIS_TEXTE),
        "kpi_label":   ps("kpi_label",
            fontName="Helvetica", fontSize=7.5, textColor=BLANC,
            alignment=TA_CENTER),
        # MISSION 2 FIX : kpi_value textColor = BLANC (etait BLEU_CIEL)
        "kpi_value":   ps("kpi_value",
            fontName="Helvetica-Bold", fontSize=16, textColor=BLANC,
            alignment=TA_CENTER),
        "disclaimer":  ps("disclaimer",
            fontName="Helvetica-Oblique", fontSize=6.5, textColor=GRIS_TEXTE,
            alignment=TA_CENTER, leading=10),
        "perf_th":     ps("perf_th",
            fontName="Helvetica-Bold", fontSize=7.5, textColor=BLANC,
            alignment=TA_RIGHT),
        "perf_th_l":   ps("perf_th_l",
            fontName="Helvetica-Bold", fontSize=7.5, textColor=BLANC,
            alignment=TA_LEFT),
        "perf_td":     ps("perf_td",
            fontName="Helvetica", fontSize=7.5, textColor=BLEU_MARINE,
            alignment=TA_RIGHT),
        "perf_td_l":   ps("perf_td_l",
            fontName="Helvetica", fontSize=7.5, textColor=BLEU_MARINE,
            alignment=TA_LEFT),
        "perf_pos":    ps("perf_pos",
            fontName="Helvetica-Bold", fontSize=7.5,
            textColor=BLEU_CIEL, alignment=TA_RIGHT),
        "perf_neg":    ps("perf_neg",
            fontName="Helvetica-Bold", fontSize=7.5,
            textColor=HexColor("#8B2020"), alignment=TA_RIGHT),
        "chart_caption": ps("chart_caption",
            fontName="Helvetica-Oblique", fontSize=7, textColor=GRIS_TEXTE,
            alignment=TA_CENTER, spaceAfter=6),
    }


# ---------------------------------------------------------------------------
# PAGE DE GARDE
# ---------------------------------------------------------------------------

def _page_garde(styles: dict, mode_comex: bool, kpis: dict,
                fonds_perimetre: Optional[list] = None) -> list:
    elements = []
    today_str     = date.today().strftime("%d %B %Y")
    perimetre_str = (
        ", ".join(fonds_perimetre) if fonds_perimetre else "Tous les fonds"
    )

    cover_data = [[Paragraph(
        f"<font color='#00A8E1' size='7.5'>CONFIDENTIEL"
        f"{'  |  MODE COMEX ACTIVE' if mode_comex else ''}</font><br/><br/>"
        f"<font color='white' size='22'><b>Executive Report</b></font><br/>"
        f"<font color='white' size='16'>Pipeline &amp; Performance</font><br/><br/>"
        f"<font color='#00A8E1' size='9'>Asset Management Division</font><br/>"
        f"<font color='white' size='7.5'>Perimetre : {perimetre_str}</font><br/><br/>"
        f"<font color='white' size='7.5'>{today_str}</font>",
        styles["cover_inner"]
    )]]
    cover_tbl = Table(cover_data, colWidths=[USABLE_W])
    cover_tbl.setStyle(TableStyle([
        ("BACKGROUND",   (0, 0), (-1, -1), BLEU_MARINE),
        ("TOPPADDING",   (0, 0), (-1, -1), 40),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 50),
        ("LEFTPADDING",  (0, 0), (-1, -1), 28),
        ("RIGHTPADDING", (0, 0), (-1, -1), 18),
    ]))
    elements.append(cover_tbl)
    elements.append(Spacer(1, 3))
    elements.append(ColorRect(USABLE_W, 4, BLEU_CIEL))
    elements.append(Spacer(1, 14))

    # Cartes KPI — kpi_value textColor=BLANC (MISSION 2 FIX)
    kpi_items = [
        ("AUM Finance Total", fmt_aum(kpis.get("total_funded", 0))),
        ("Pipeline Actif",    fmt_aum(kpis.get("pipeline_actif", 0))),
        ("Taux Conversion",   f"{kpis.get('taux_conversion', 0):.1f}%"),
        ("Deals Actifs",      str(kpis.get("nb_deals_actifs", 0))),
    ]
    col_w = USABLE_W / 4

    kpi_tbl = Table(
        [[Paragraph(i[0], styles["kpi_label"]) for i in kpi_items],
         [Paragraph(i[1], styles["kpi_value"]) for i in kpi_items]],
        colWidths=[col_w] * 4,
        rowHeights=[0.95 * cm, 1.15 * cm]
    )
    kpi_tbl.setStyle(TableStyle([
        ("BACKGROUND",   (0, 0), (-1, -1), BLEU_MARINE),
        ("VALIGN",       (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",   (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 7),
        ("LEFTPADDING",  (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("LINEAFTER",    (0, 0), (2, -1),  0.5, BLEU_CIEL),
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
# SECTION GRAPHIQUES ANALYTIQUES (empiles verticalement — MISSION 2)
# Ordre : 1. Type Client | 2. Region | 3. Top 10 Deals
# ---------------------------------------------------------------------------

def _section_charts(
    kpis: dict,
    aum_by_region: dict,
    top_deals_safe: list,
    styles: dict,
    mode_comex: bool,
) -> list:
    """
    Graphiques empiles verticalement, largeur CHART_W = USABLE_W * 0.8.
    Centres dans la page via un tableau de centrage.
    """
    elements = []
    elements.append(Paragraph("Analyse du Pipeline", styles["section"]))
    elements.append(ColorRect(USABLE_W, 2, BLEU_CIEL))
    elements.append(Spacer(1, 12))

    # Helper : centre un graphique BytesIO
    def _add_chart(buf: io.BytesIO, aspect: float, caption: str):
        w = CHART_W
        h = w * aspect
        img = Image(buf, width=w, height=h)
        # Tableau de centrage (padding lateral = (USABLE_W - CHART_W) / 2)
        pad = (USABLE_W - CHART_W) / 2
        centering = Table([[img]], colWidths=[w])
        centering.setStyle(TableStyle([
            ("ALIGN",       (0, 0), (-1, -1), "CENTER"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING",(0, 0), (-1, -1), 0),
            ("TOPPADDING",  (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING",(0,0), (-1, -1), 0),
        ]))
        elements.append(centering)
        elements.append(Paragraph(caption, styles["chart_caption"]))
        elements.append(Spacer(1, 10))
        elements.append(ColorRect(USABLE_W * 0.4, 1, GRIS_CLAIR))
        elements.append(Spacer(1, 12))

    # 1. Type de Client
    buf_type = _chart_pie_aum_by_type(kpis.get("aum_by_type", {}))
    _add_chart(buf_type, 0.70,
               "Graphique 1 — Repartition AUM Finance par Type de Client")

    # 2. Region Geographique
    buf_reg = _chart_pie_aum_by_region(aum_by_region)
    _add_chart(buf_reg, 0.70,
               "Graphique 2 — Repartition AUM Finance par Region Geographique")

    # 3. Top 10 Deals
    n_deals = len(top_deals_safe[:10])
    aspect  = max(0.40, n_deals * 0.055 + 0.15)
    buf_top = _chart_top10_deals(top_deals_safe, mode_comex)
    _add_chart(buf_top, aspect,
               "Graphique 3 — Top 10 Deals par AUM Finance")

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
    elements.append(ColorRect(USABLE_W, 2, BLEU_CIEL))
    elements.append(Spacer(1, 8))

    active = pipeline_df[pipeline_df["statut"].isin(
        ["Prospect", "Initial Pitch", "Due Diligence", "Soft Commit"]
    )].copy()

    if active.empty:
        elements.append(Paragraph("Aucun deal actif dans le perimetre.", styles["body"]))
        return elements

    if mode_comex:
        active["nom_client"] = active.apply(
            lambda r: f"{r['type_client']} - {r['region']}", axis=1
        )

    # Ratios colonnes — somme exactement = 1.0
    ratios     = [0.23, 0.16, 0.13, 0.13, 0.13, 0.13, 0.09]
    col_widths = [USABLE_W * r for r in ratios]
    headers    = ["Client", "Fonds", "Statut",
                  "AUM Cible", "AUM Revise", "Prochaine Action", "Commercial"]

    rows = [[Paragraph(h, styles["th"]) for h in headers]]
    today = date.today()

    for _, row in active.iterrows():
        nad = row.get("next_action_date")
        if isinstance(nad, date):
            nad_str = f"[!] {nad.isoformat()}" if nad < today else nad.isoformat()
        else:
            nad_str = "—"

        rows.append([
            Paragraph(str(row.get("nom_client", ""))[:28],    styles["td"]),
            Paragraph(str(row.get("fonds", "")),              styles["td"]),
            Paragraph(str(row.get("statut", "")),             styles["td"]),
            Paragraph(fmt_aum(float(row.get("target_aum_initial", 0) or 0)),
                      styles["td"]),
            Paragraph(fmt_aum(float(row.get("revised_aum", 0) or 0)),
                      styles["td"]),
            Paragraph(nad_str,                                styles["td"]),
            Paragraph(str(row.get("sales_owner", ""))[:16],  styles["td"]),
        ])

    tbl = Table(rows, colWidths=col_widths, repeatRows=1)
    tbl.setStyle(TableStyle([
        ("BACKGROUND",   (0, 0), (-1, 0),  BLEU_MARINE),
        ("LINEBELOW",    (0, 0), (-1, 0),  1.5, BLEU_CIEL),
        ("GRID",         (0, 0), (-1, -1), 0.35, GRIS_CLAIR),
        ("VALIGN",       (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",   (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 4),
        ("LEFTPADDING",  (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        *[("BACKGROUND", (0, i), (-1, i), HexColor("#F5F8FC"))
          for i in range(2, len(rows), 2)],
    ]))
    elements.append(tbl)

    # Tableau Lost / Paused
    lost_paused = pipeline_df[
        pipeline_df["statut"].isin(["Lost", "Paused"])
    ].copy()
    if not lost_paused.empty:
        elements.append(Spacer(1, 14))
        elements.append(Paragraph("Deals Perdus / En Pause", styles["section"]))
        elements.append(ColorRect(USABLE_W, 2, GRIS_CLAIR))
        elements.append(Spacer(1, 8))

        if mode_comex:
            lost_paused["nom_client"] = lost_paused.apply(
                lambda r: f"{r['type_client']} - {r['region']}", axis=1
            )

        lp_ratios  = [0.27, 0.19, 0.13, 0.22, 0.19]
        lp_col_w   = [USABLE_W * r for r in lp_ratios]
        lp_headers = ["Client", "Fonds", "Statut", "Raison", "Concurrent"]
        lp_rows    = [[Paragraph(h, styles["th"]) for h in lp_headers]]

        for _, row in lost_paused.iterrows():
            lp_rows.append([
                Paragraph(str(row.get("nom_client",""))[:26],        styles["td_grey"]),
                Paragraph(str(row.get("fonds","")),                  styles["td_grey"]),
                Paragraph(str(row.get("statut","")),                 styles["td_grey"]),
                Paragraph(str(row.get("raison_perte","") or "—"),    styles["td_grey"]),
                Paragraph(str(row.get("concurrent_choisi","") or "—"), styles["td_grey"]),
            ])

        lp_tbl = Table(lp_rows, colWidths=lp_col_w, repeatRows=1)
        lp_tbl.setStyle(TableStyle([
            ("BACKGROUND",   (0, 0), (-1, 0),  HexColor("#888888")),
            ("LINEBELOW",    (0, 0), (-1, 0),  1.0, GRIS_CLAIR),
            ("GRID",         (0, 0), (-1, -1), 0.35, GRIS_CLAIR),
            ("VALIGN",       (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING",   (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING",(0, 0), (-1, -1), 4),
            ("LEFTPADDING",  (0, 0), (-1, -1), 5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ]))
        elements.append(lp_tbl)

    return elements


# ---------------------------------------------------------------------------
# SECTION PERFORMANCE (optionnelle, filtrée sur fonds_perimetre)
# ---------------------------------------------------------------------------

def _section_performance(
    perf_data: pd.DataFrame,
    nav_base100_df: Optional[pd.DataFrame],
    styles: dict,
    fonds_perimetre: Optional[list] = None,
) -> list:
    """
    Page Performance. Si fonds_perimetre est fourni, filtre perf_data
    et nav_base100_df sur ces fonds uniquement.
    """
    elements = []
    elements.append(PageBreak())
    elements.append(Paragraph("Analyse de Performance — NAV", styles["section"]))
    elements.append(ColorRect(USABLE_W, 2, BLEU_CIEL))
    elements.append(Spacer(1, 10))

    # Filtrage sur le perimetre
    pf = perf_data.copy() if perf_data is not None else pd.DataFrame()
    nb = nav_base100_df.copy() if nav_base100_df is not None else None

    if fonds_perimetre:
        if not pf.empty and "Fonds" in pf.columns:
            pf = pf[pf["Fonds"].isin(fonds_perimetre)]
        if nb is not None and not nb.empty:
            cols_keep = [c for c in nb.columns if c in fonds_perimetre]
            nb = nb[cols_keep] if cols_keep else pd.DataFrame()

    # Graphique NAV Base 100
    buf_nav = _chart_nav_base100(nb)
    img_h   = USABLE_W * 0.36
    elements.append(Image(buf_nav, width=USABLE_W, height=img_h))
    elements.append(Spacer(1, 14))

    if pf.empty:
        elements.append(Paragraph(
            "Aucune donnee de performance disponible pour ce perimetre.",
            styles["body"]
        ))
        return elements

    # Tableau des performances
    elements.append(Paragraph("Tableau des Performances", styles["section"]))
    elements.append(ColorRect(USABLE_W, 2, BLEU_CIEL))
    elements.append(Spacer(1, 8))

    perf_col_names = {
        "Perf 1M (%)", "Perf YTD (%)", "Perf Periode (%)"
    }
    col_order  = ["Fonds", "NAV Derniere", "Base 100 Actuel",
                  "Perf 1M (%)", "Perf YTD (%)", "Perf Periode (%)"]
    available  = [c for c in col_order if c in pf.columns]
    n          = len(available)

    if n == 0:
        elements.append(Paragraph("Colonnes manquantes dans les donnees de performance.",
                                  styles["body"]))
        return elements

    fonds_r  = 0.28
    rest_r   = (1.0 - fonds_r) / max(n - 1, 1)
    col_widths = [USABLE_W * (fonds_r if i == 0 else rest_r)
                  for i in range(n)]

    h_row  = [Paragraph(c, styles["perf_th_l"] if i == 0 else styles["perf_th"])
              for i, c in enumerate(available)]
    rows   = [h_row]

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
        ("BACKGROUND",   (0, 0), (-1, 0),  BLEU_MARINE),
        ("LINEBELOW",    (0, 0), (-1, 0),  1.5, BLEU_CIEL),
        ("GRID",         (0, 0), (-1, -1), 0.35, GRIS_CLAIR),
        ("VALIGN",       (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",   (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 5),
        ("LEFTPADDING",  (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        *[("BACKGROUND", (0, i), (-1, i), HexColor("#F5F8FC"))
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
        canvas.setFillColor(BLEU_MARINE)
        canvas.rect(0, 0, PAGE_W, 1.0 * cm, fill=1, stroke=0)
        canvas.setFont("Helvetica", 6.5)
        canvas.setFillColor(GRIS_CLAIR)
        canvas.drawString(2 * cm, 0.35 * cm,
                          "Executive Report — Asset Management Division")
        canvas.drawRightString(PAGE_W - 2 * cm, 0.35 * cm, f"Page {doc.page}")
        canvas.setStrokeColor(BLEU_CIEL)
        canvas.setLineWidth(2.5)
        canvas.line(0, PAGE_H - 0.45 * cm, PAGE_W, PAGE_H - 0.45 * cm)
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(BLEU_MARINE)
        canvas.drawRightString(PAGE_W - 2 * cm, PAGE_H - 0.32 * cm, "CONFIDENTIEL")
    canvas.restoreState()


# ---------------------------------------------------------------------------
# FONCTION PRINCIPALE generate_pdf
# ---------------------------------------------------------------------------

def generate_pdf(
    pipeline_df: pd.DataFrame,
    kpis: dict,
    aum_by_region: Optional[dict] = None,
    mode_comex: bool = False,
    perf_data: Optional[pd.DataFrame] = None,
    nav_base100_df: Optional[pd.DataFrame] = None,
    fonds_perimetre: Optional[list] = None,
) -> bytes:
    """
    Genere le rapport PDF Pitchbook complet.

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
            {**d, "nom_client": f"{d['type_client']} - {d['region']}"}
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
        leftMargin=MARGIN_H, rightMargin=MARGIN_H,
        topMargin=MARGIN_V,  bottomMargin=MARGIN_V,
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
    if perf_data is not None and not perf_data.empty:
        elements += _section_performance(
            perf_data, nav_base100_df, styles, fonds_perimetre
        )

    doc.build(elements,
              onFirstPage=_header_footer,
              onLaterPages=_header_footer)

    result = pdf_buf.getvalue()
    pdf_buf.close()
    return result
