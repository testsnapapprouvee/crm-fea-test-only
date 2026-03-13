# =============================================================================
# pdf_generator.py — Rapport PDF style Pitchbook — Edition Enterprise
# Charte Amundi : Marine #002D54 | Ciel #00A8E1
# Staff Engineer refactoring : marges strictes, page Performance optionnelle
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
from reportlab.lib.units import cm, mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.colors import HexColor, white
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    Image, PageBreak, KeepTogether
)
from reportlab.platypus.flowables import Flowable


# ---------------------------------------------------------------------------
# CONSTANTES CHARTE ET MISE EN PAGE
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

PAGE_W, PAGE_H = A4          # 595.28 x 841.89 pts
MARGIN_H       = 2 * cm      # marges gauche/droite
MARGIN_V       = 1.8 * cm    # marges haut/bas
USABLE_W       = PAGE_W - 2 * MARGIN_H   # 481.89 pts — largeur utilisable stricte

# Palette NAV identique a app.py
NAV_PALETTE = [HEX_CIEL, HEX_MARINE, "#1A6B9A", "#7BC8E8", "#003F7A",
               "#5BA3C9", "#2C8FBF", "#004F8C", "#A8D8EE", "#003060"]

MPL_PARAMS = {
    "font.family":       "DejaVu Sans",
    "axes.spines.top":   False,
    "axes.spines.right": False,
    "axes.labelcolor":   HEX_MARINE,
    "xtick.color":       HEX_MARINE,
    "ytick.color":       HEX_MARINE,
    "figure.facecolor":  "white",
    "axes.facecolor":    "white",
    "grid.color":        HEX_GRIS,
    "grid.linewidth":    0.5,
}


# ---------------------------------------------------------------------------
# HELPER : ANONYMISATION MODE COMEX
# ---------------------------------------------------------------------------

def anonymize_df(df: pd.DataFrame) -> pd.DataFrame:
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


# ---------------------------------------------------------------------------
# FLOWABLES PERSONNALISES
# ---------------------------------------------------------------------------

class ColorRect(Flowable):
    """Rectangle colore (separateurs, bandeaux de section)."""
    def __init__(self, width, height, color):
        super().__init__()
        self.width  = width
        self.height = height
        self.color  = color

    def draw(self):
        self.canv.setFillColor(self.color)
        self.canv.rect(0, 0, self.width, self.height, fill=1, stroke=0)


# ---------------------------------------------------------------------------
# STYLES REPORTLAB
# ---------------------------------------------------------------------------

def _build_styles() -> dict:
    s = {}

    def ps(name, **kw):
        return ParagraphStyle(name, **kw)

    s["cover_inner"] = ps("cover_inner",
        fontName="Helvetica", fontSize=9, textColor=BLANC, leading=16)
    s["section"]     = ps("section",
        fontName="Helvetica-Bold", fontSize=13, textColor=BLEU_MARINE,
        spaceBefore=14, spaceAfter=7)
    s["body"]        = ps("body",
        fontName="Helvetica", fontSize=8.5, textColor=GRIS_TEXTE,
        spaceAfter=4, leading=13)
    s["th"]          = ps("th",
        fontName="Helvetica-Bold", fontSize=7.5, textColor=BLANC)
    s["td"]          = ps("td",
        fontName="Helvetica", fontSize=7.5, textColor=BLEU_MARINE)
    s["td_grey"]     = ps("td_grey",
        fontName="Helvetica", fontSize=7.5, textColor=GRIS_TEXTE)
    s["kpi_label"]   = ps("kpi_label",
        fontName="Helvetica", fontSize=7.5, textColor=BLANC, alignment=TA_CENTER)
    s["kpi_value"]   = ps("kpi_value",
        fontName="Helvetica-Bold", fontSize=16, textColor=BLEU_CIEL,
        alignment=TA_CENTER)
    s["footer"]      = ps("footer",
        fontName="Helvetica", fontSize=6.5, textColor=GRIS_TEXTE, alignment=TA_CENTER)
    s["disclaimer"]  = ps("disclaimer",
        fontName="Helvetica-Oblique", fontSize=6.5, textColor=GRIS_TEXTE,
        alignment=TA_CENTER, leading=10)
    s["perf_th"]     = ps("perf_th",
        fontName="Helvetica-Bold", fontSize=7.5, textColor=BLANC, alignment=TA_RIGHT)
    s["perf_th_l"]   = ps("perf_th_l",
        fontName="Helvetica-Bold", fontSize=7.5, textColor=BLANC, alignment=TA_LEFT)
    s["perf_td"]     = ps("perf_td",
        fontName="Helvetica", fontSize=7.5, textColor=BLEU_MARINE, alignment=TA_RIGHT)
    s["perf_td_l"]   = ps("perf_td_l",
        fontName="Helvetica", fontSize=7.5, textColor=BLEU_MARINE, alignment=TA_LEFT)
    s["perf_pos"]    = ps("perf_pos",
        fontName="Helvetica-Bold", fontSize=7.5, textColor=BLEU_CIEL, alignment=TA_RIGHT)
    s["perf_neg"]    = ps("perf_neg",
        fontName="Helvetica-Bold", fontSize=7.5, textColor=HexColor("#8B2020"),
        alignment=TA_RIGHT)
    return s


# ---------------------------------------------------------------------------
# GRAPHIQUES MATPLOTLIB -> BYTESIO
# ---------------------------------------------------------------------------

def _chart_pie_aum_by_type(aum_by_type: dict) -> io.BytesIO:
    plt.rcParams.update(MPL_PARAMS)
    fig, ax = plt.subplots(figsize=(5.2, 4.5))

    if not aum_by_type or sum(aum_by_type.values()) == 0:
        ax.text(0.5, 0.5, "Aucune donnee Funded",
                ha="center", va="center", color=HEX_MARINE)
        ax.axis("off")
    else:
        lbls   = list(aum_by_type.keys())
        vals   = list(aum_by_type.values())
        colors = [NAV_PALETTE[i % len(NAV_PALETTE)] for i in range(len(lbls))]

        _, _, autotexts = ax.pie(
            vals, colors=colors, autopct="%1.1f%%", startangle=90,
            pctdistance=0.74,
            wedgeprops={"width": 0.54, "edgecolor": "white", "linewidth": 1.5}
        )
        for at in autotexts:
            at.set_fontsize(8.5); at.set_color("white"); at.set_fontweight("bold")

        total = sum(vals)
        ax.text(0, 0.08, fmt_aum(total), ha="center", va="center",
                fontsize=10, fontweight="bold", color=HEX_MARINE)
        ax.text(0, -0.18, "AUM Finance", ha="center", va="center",
                fontsize=7.5, color=HEX_GTXT)

        patches = [mpatches.Patch(color=colors[i],
                   label=f"{lbls[i]}: {fmt_aum(vals[i])}")
                   for i in range(len(lbls))]
        ax.legend(handles=patches, loc="lower center",
                  bbox_to_anchor=(0.5, -0.22), ncol=2,
                  fontsize=7, frameon=False, labelcolor=HEX_MARINE)
        ax.set_title("Repartition AUM Funded par Type Client",
                     fontsize=9.5, fontweight="bold", color=HEX_MARINE, pad=8)

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150,
                bbox_inches="tight", facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return buf


def _chart_top10_deals(top_deals: list, mode_comex: bool) -> io.BytesIO:
    plt.rcParams.update(MPL_PARAMS)
    fig, ax = plt.subplots(
        figsize=(7.2, max(3.2, len(top_deals) * 0.44 + 0.8))
    )

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

        bars = ax.barh(range(len(labels)), values,
                       color=HEX_CIEL, edgecolor="white", height=0.6)

        max_v = max(values) if values else 1
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
        ax.set_xlim(0, max_v * 1.2)
        ax.grid(axis="x", alpha=0.3)
        ax.set_title("Top 10 Deals — AUM Finance",
                     fontsize=9.5, fontweight="bold", color=HEX_MARINE, pad=8)

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150,
                bbox_inches="tight", facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return buf


def _chart_nav_base100(nav_base100_df: pd.DataFrame) -> io.BytesIO:
    """Genere le graphique NAV Base 100 pour l'integration PDF."""
    plt.rcParams.update(MPL_PARAMS)
    fig, ax = plt.subplots(figsize=(12, 4.8))

    if nav_base100_df is None or nav_base100_df.empty:
        ax.text(0.5, 0.5, "Donnees NAV non disponibles",
                ha="center", va="center", color=HEX_MARINE, fontsize=11)
        ax.axis("off")
    else:
        fonds_list = [c for c in nav_base100_df.columns]
        for i, fonds in enumerate(fonds_list):
            series = nav_base100_df[fonds].dropna()
            if series.empty:
                continue
            color    = NAV_PALETTE[i % len(NAV_PALETTE)]
            line_sty = "-" if i % 2 == 0 else "--"
            line_w   = 1.8 if i < 3 else 1.4

            if len(series) >= 2:
                ax.plot(series.index, series.values,
                        label=fonds, color=color,
                        linewidth=line_w, linestyle=line_sty, alpha=0.9)
            else:
                # Un seul point : marqueur uniquement
                ax.scatter(series.index, series.values,
                           color=color, s=60, label=fonds, zorder=5)

            if not series.empty:
                ax.scatter([series.index[-1]], [series.values[-1]],
                           color=color, s=35, zorder=6)

        ax.axhline(100, color=HEX_GRIS, linewidth=0.8, linestyle=":")
        ax.set_ylabel("NAV (Base 100)", fontsize=8.5, color=HEX_MARINE)
        ax.tick_params(colors=HEX_MARINE, labelsize=7.5)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.grid(axis="y", alpha=0.25, color=HEX_GRIS)
        ax.legend(fontsize=7.5, frameon=True, framealpha=0.9,
                  edgecolor=HEX_GRIS, labelcolor=HEX_MARINE,
                  loc="upper left", ncol=min(3, len(fonds_list)))
        ax.set_title("Evolution NAV — Base 100",
                     fontsize=10, fontweight="bold", color=HEX_MARINE, pad=8)
        plt.xticks(rotation=18, ha="right", fontsize=7.5)

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150,
                bbox_inches="tight", facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return buf


# ---------------------------------------------------------------------------
# PAGE DE GARDE
# ---------------------------------------------------------------------------

def _page_garde(styles, mode_comex: bool, kpis: dict) -> list:
    elements = []
    today_str = date.today().strftime("%d %B %Y")

    # Bandeau couverture
    cover_data = [[Paragraph(
        f"<font color='#00A8E1' size='8'>CONFIDENTIEL"
        f"{'  |  MODE COMEX ACTIVE' if mode_comex else ''}</font><br/><br/>"
        f"<font color='white' size='24'><b>Executive Report</b></font><br/>"
        f"<font color='white' size='18'>Pipeline &amp; Performance</font><br/><br/>"
        f"<font color='#00A8E1' size='10'>Asset Management Division</font><br/><br/>"
        f"<font color='white' size='8'>{today_str}</font>",
        styles["cover_inner"]
    )]]
    cover_tbl = Table(cover_data, colWidths=[USABLE_W])
    cover_tbl.setStyle(TableStyle([
        ("BACKGROUND",   (0, 0), (-1, -1), BLEU_MARINE),
        ("TOPPADDING",   (0, 0), (-1, -1), 42),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 52),
        ("LEFTPADDING",  (0, 0), (-1, -1), 30),
        ("RIGHTPADDING", (0, 0), (-1, -1), 20),
    ]))
    elements.append(cover_tbl)
    elements.append(Spacer(1, 3))
    elements.append(ColorRect(USABLE_W, 4, BLEU_CIEL))
    elements.append(Spacer(1, 14))

    # Cartes KPI sur la couverture — 4 colonnes egales
    kpi_items = [
        ("AUM Finance Total",  fmt_aum(kpis.get("total_funded", 0))),
        ("Pipeline Actif",     fmt_aum(kpis.get("pipeline_actif", 0))),
        ("Taux Conversion",    f"{kpis.get('taux_conversion', 0):.1f}%"),
        ("Deals Actifs",       str(kpis.get("nb_deals_actifs", 0))),
    ]
    col_w = USABLE_W / 4
    kpi_labels_row = [[Paragraph(i[0], styles["kpi_label"]) for i in kpi_items]]
    kpi_values_row = [[Paragraph(i[1], styles["kpi_value"]) for i in kpi_items]]

    kpi_tbl = Table(
        kpi_labels_row + kpi_values_row,
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
        "Ce document est strictement confidentiel et destine exclusivement a l'usage interne. "
        "Toute reproduction ou diffusion externe est strictement interdite. "
        "Les performances passees ne prejudgent pas des performances futures."
    )
    if mode_comex:
        disclaimer += " Mode Comex : les noms clients ont ete anonymises."
    elements.append(Paragraph(disclaimer, styles["disclaimer"]))
    elements.append(PageBreak())
    return elements


# ---------------------------------------------------------------------------
# TABLEAU PIPELINE ACTIF
# ---------------------------------------------------------------------------

def _section_pipeline(pipeline_df: pd.DataFrame,
                      styles: dict, mode_comex: bool) -> list:
    elements = []
    elements.append(Paragraph("Pipeline Actif — Recapitulatif", styles["section"]))
    elements.append(ColorRect(USABLE_W, 2, BLEU_CIEL))
    elements.append(Spacer(1, 8))

    active = pipeline_df[pipeline_df["statut"].isin(
        ["Prospect", "Initial Pitch", "Due Diligence", "Soft Commit"]
    )].copy()

    if active.empty:
        elements.append(Paragraph("Aucun deal actif.", styles["body"]))
        return elements

    if mode_comex:
        active["nom_client"] = active.apply(
            lambda r: f"{r['type_client']} - {r['region']}", axis=1
        )

    # Colonnes et proportions : somme = 1.0 -> USABLE_W
    ratios    = [0.24, 0.17, 0.13, 0.13, 0.13, 0.12, 0.08]
    col_widths = [USABLE_W * r for r in ratios]
    headers   = ["Client", "Fonds", "Statut",
                 "AUM Cible", "AUM Revise", "Prochaine Action", "Commercial"]

    header_row = [Paragraph(h, styles["th"]) for h in headers]
    rows       = [header_row]
    today      = date.today()

    for _, row in active.iterrows():
        nad = row.get("next_action_date")
        if isinstance(nad, date):
            if nad < today:
                nad_str = f"[RETARD] {nad.isoformat()}"
            else:
                nad_str = nad.isoformat()
        else:
            nad_str = "—"

        rows.append([
            Paragraph(str(row.get("nom_client", ""))[:28],    styles["td"]),
            Paragraph(str(row.get("fonds", "")),              styles["td"]),
            Paragraph(str(row.get("statut", "")),             styles["td"]),
            Paragraph(fmt_aum(float(row.get("target_aum_initial", 0) or 0)), styles["td"]),
            Paragraph(fmt_aum(float(row.get("revised_aum", 0) or 0)),        styles["td"]),
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
    return elements


# ---------------------------------------------------------------------------
# PAGE PERFORMANCE (optionnelle)
# ---------------------------------------------------------------------------

def _section_performance(
    perf_data: pd.DataFrame,
    nav_base100_df: Optional[pd.DataFrame],
    styles: dict
) -> list:
    """
    Genere la page Performance & NAV.
    perf_data : DataFrame avec colonnes Fonds, NAV Derniere,
                Base 100 Actuel, Perf 1M (%), Perf YTD (%), Perf Periode (%)
    nav_base100_df : DataFrame pivot (index=Date, cols=Fonds) Base100
    """
    elements = []
    elements.append(PageBreak())
    elements.append(Paragraph("Analyse de Performance — NAV", styles["section"]))
    elements.append(ColorRect(USABLE_W, 2, BLEU_CIEL))
    elements.append(Spacer(1, 10))

    # --- Graphique NAV Base 100 ---
    buf_nav = _chart_nav_base100(nav_base100_df)
    img_h   = USABLE_W * 0.40
    elements.append(Image(buf_nav, width=USABLE_W, height=img_h))
    elements.append(Spacer(1, 14))

    # --- Tableau des performances ---
    elements.append(Paragraph("Tableau des Performances", styles["section"]))
    elements.append(ColorRect(USABLE_W, 2, BLEU_CIEL))
    elements.append(Spacer(1, 8))

    if perf_data is None or perf_data.empty:
        elements.append(Paragraph("Aucune donnee de performance disponible.",
                                  styles["body"]))
        return elements

    # Colonnes et largeurs — somme = USABLE_W
    perf_cols = ["Fonds", "NAV Derniere", "Base 100 Actuel",
                 "Perf 1M (%)", "Perf YTD (%)", "Perf Periode (%)"]
    # Colonnes disponibles dans le df fourni
    available = [c for c in perf_cols if c in perf_data.columns]
    n         = len(available)

    # Largeur fonds plus large, reste equitable
    if n == 0:
        elements.append(Paragraph("Donnees incompletes.", styles["body"]))
        return elements

    fonds_ratio = 0.28
    rest_ratio  = (1 - fonds_ratio) / max(n - 1, 1)
    ratios      = [fonds_ratio] + [rest_ratio] * (n - 1)
    col_widths  = [USABLE_W * r for r in ratios]

    # Header
    h_styles = [styles["perf_th_l"]] + [styles["perf_th"]] * (n - 1)
    header_row = [Paragraph(c, h_styles[i]) for i, c in enumerate(available)]
    rows = [header_row]

    def _fmt_perf(val):
        """Formate une valeur de perf avec signe."""
        if val is None or (isinstance(val, float) and np.isnan(val)):
            return "—"
        sign = "+" if val > 0 else ""
        return f"{sign}{val:.2f}%"

    def _perf_style(val, styles_dict):
        if val is None or (isinstance(val, float) and np.isnan(val)):
            return styles_dict["perf_td"]
        return styles_dict["perf_pos"] if val >= 0 else styles_dict["perf_neg"]

    perf_col_names = {"Perf 1M (%)", "Perf YTD (%)", "Perf Periode (%)"}

    for i, (_, row) in enumerate(perf_data.iterrows()):
        data_row = []
        for j, col in enumerate(available):
            val = row.get(col)
            if col == "Fonds":
                data_row.append(Paragraph(str(val), styles["perf_td_l"]))
            elif col in perf_col_names:
                fval    = float(val) if val is not None and val == val else None
                s       = _perf_style(fval, styles)
                data_row.append(Paragraph(_fmt_perf(fval), s))
            elif col in ("NAV Derniere", "Base 100 Actuel"):
                try:
                    fval = float(val)
                    txt  = f"{fval:.4f}" if col == "NAV Derniere" else f"{fval:.2f}"
                except Exception:
                    txt = "—"
                data_row.append(Paragraph(txt, styles["perf_td"]))
            else:
                data_row.append(Paragraph(str(val) if val is not None else "—",
                                          styles["perf_td"]))
        rows.append(data_row)

    perf_tbl = Table(rows, colWidths=col_widths, repeatRows=1)
    tbl_style = TableStyle([
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
    ])
    perf_tbl.setStyle(tbl_style)
    elements.append(perf_tbl)

    return elements


# ---------------------------------------------------------------------------
# CALLBACK EN-TETE / PIED DE PAGE
# ---------------------------------------------------------------------------

def _header_footer(canvas, doc):
    canvas.saveState()
    if doc.page > 1:
        # Bande marine en bas
        canvas.setFillColor(BLEU_MARINE)
        canvas.rect(0, 0, PAGE_W, 1.0 * cm, fill=1, stroke=0)
        canvas.setFont("Helvetica", 6.5)
        canvas.setFillColor(GRIS_CLAIR)
        canvas.drawString(2 * cm, 0.35 * cm,
                          "Executive Report — Asset Management Division")
        canvas.drawRightString(PAGE_W - 2 * cm, 0.35 * cm, f"Page {doc.page}")

        # Filet ciel en haut
        canvas.setStrokeColor(BLEU_CIEL)
        canvas.setLineWidth(2.5)
        canvas.line(0, PAGE_H - 0.45 * cm, PAGE_W, PAGE_H - 0.45 * cm)
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(BLEU_MARINE)
        canvas.drawRightString(PAGE_W - 2 * cm, PAGE_H - 0.32 * cm, "CONFIDENTIEL")
    canvas.restoreState()


# ---------------------------------------------------------------------------
# FONCTION PRINCIPALE
# ---------------------------------------------------------------------------

def generate_pdf(
    pipeline_df: pd.DataFrame,
    kpis: dict,
    mode_comex: bool = False,
    perf_data: Optional[pd.DataFrame] = None,
    nav_base100_df: Optional[pd.DataFrame] = None,
) -> bytes:
    """
    Genere le rapport PDF complet.

    Args:
        pipeline_df    : DataFrame pipeline (enrichi clients).
        kpis           : Dict KPIs depuis database.get_kpis().
        mode_comex     : Anonymisation totale des noms clients.
        perf_data      : DataFrame performances NAV (optionnel).
                         Colonnes attendues : Fonds, NAV Derniere,
                         Base 100 Actuel, Perf 1M (%), Perf YTD (%),
                         Perf Periode (%).
        nav_base100_df : DataFrame pivot Base100 pour le graphique (optionnel).

    Returns:
        bytes du PDF pret au telechargement.
    """
    # --- Anonymisation Mode Comex (etanche : aucun nom ne survit) ---
    if mode_comex:
        pipeline_df = anonymize_df(pipeline_df)
        top_deals_safe = []
        for d in kpis.get("top_deals", []):
            d2 = dict(d)
            d2["nom_client"] = f"{d2['type_client']} - {d2['region']}"
            top_deals_safe.append(d2)
        kpis = {**kpis, "top_deals": top_deals_safe}
    else:
        top_deals_safe = kpis.get("top_deals", [])

    styles   = _build_styles()
    pdf_buf  = io.BytesIO()

    doc = SimpleDocTemplate(
        pdf_buf,
        pagesize=A4,
        leftMargin=MARGIN_H, rightMargin=MARGIN_H,
        topMargin=MARGIN_V,  bottomMargin=MARGIN_V,
        title="Executive Report — Asset Management",
        author="Asset Management Division",
    )

    elements = []

    # PAGE 1 : Couverture
    elements += _page_garde(styles, mode_comex, kpis)

    # PAGE 2 : Graphiques AUM
    elements.append(Paragraph("Analyse du Pipeline", styles["section"]))
    elements.append(ColorRect(USABLE_W, 2, BLEU_CIEL))
    elements.append(Spacer(1, 10))

    chart_w = (USABLE_W - 6) / 2   # 2 colonnes avec petit gap
    chart_h = chart_w * 0.82

    buf_pie = _chart_pie_aum_by_type(kpis.get("aum_by_type", {}))
    buf_bar = _chart_top10_deals(top_deals_safe, mode_comex)

    img_pie = Image(buf_pie, width=chart_w, height=chart_h)
    img_bar = Image(buf_bar, width=chart_w, height=chart_h)

    charts_tbl = Table(
        [[img_pie, img_bar]],
        colWidths=[chart_w + 3, chart_w + 3]
    )
    charts_tbl.setStyle(TableStyle([
        ("ALIGN",        (0, 0), (-1, -1), "CENTER"),
        ("VALIGN",       (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING",  (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
    ]))
    elements.append(charts_tbl)
    elements.append(Spacer(1, 12))

    # PAGE 3 : Tableau pipeline
    elements.append(PageBreak())
    elements += _section_pipeline(pipeline_df, styles, mode_comex)

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

        lp_ratios    = [0.28, 0.19, 0.14, 0.22, 0.17]
        lp_col_w     = [USABLE_W * r for r in lp_ratios]
        lp_headers   = ["Client", "Fonds", "Statut", "Raison", "Concurrent"]
        lp_rows      = [[Paragraph(h, styles["th"]) for h in lp_headers]]

        for _, row in lost_paused.iterrows():
            lp_rows.append([
                Paragraph(str(row.get("nom_client",""))[:26], styles["td_grey"]),
                Paragraph(str(row.get("fonds","")),           styles["td_grey"]),
                Paragraph(str(row.get("statut","")),          styles["td_grey"]),
                Paragraph(str(row.get("raison_perte","") or "—"), styles["td_grey"]),
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

    # PAGE 4 (optionnelle) : Performance & NAV
    if perf_data is not None and not perf_data.empty:
        elements += _section_performance(perf_data, nav_base100_df, styles)

    doc.build(elements,
              onFirstPage=_header_footer,
              onLaterPages=_header_footer)

    pdf_bytes = pdf_buf.getvalue()
    pdf_buf.close()
    return pdf_bytes
