# =============================================================================
# pdf_generator.py — Génération du rapport PDF style Pitchbook
# Charte Amundi : Bleu Marine #002D54 | Bleu Ciel #00A8E1
# =============================================================================

import io
import os
from datetime import date
from typing import Optional

import matplotlib
matplotlib.use("Agg")  # Backend non-interactif (obligatoire pour les serveurs)
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import pandas as pd
import numpy as np

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm, mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.colors import HexColor, white, black
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    Image, HRFlowable, PageBreak, KeepTogether
)
from reportlab.platypus.flowables import Flowable


# ---------------------------------------------------------------------------
# CONSTANTES — CHARTE AMUNDI
# ---------------------------------------------------------------------------

BLEU_MARINE   = HexColor("#002D54")
BLEU_CIEL     = HexColor("#00A8E1")
BLANC         = HexColor("#FFFFFF")
GRIS_CLAIR    = HexColor("#E0E0E0")
GRIS_TEXTE    = HexColor("#555555")

HEX_MARINE    = "#002D54"
HEX_CIEL      = "#00A8E1"
HEX_GRIS      = "#E0E0E0"
HEX_GRIS_TEXT = "#555555"

PAGE_W, PAGE_H = A4  # 595.27 x 841.89 points

# Paramètres Matplotlib globaux
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
    """
    Mode Comex : remplace 'nom_client' par '{type_client} - {region}'.
    Travaille sur une COPIE — aucun nom original ne survit dans le DataFrame retourné.
    """
    df = df.copy()
    if "nom_client" in df.columns and "type_client" in df.columns and "region" in df.columns:
        df["nom_client"] = df.apply(
            lambda r: f"{r['type_client']} – {r['region']}", axis=1
        )
    return df


def fmt_aum(value: float) -> str:
    """Formate un montant AUM : 42_000_000 → '42.0M€'"""
    if value >= 1_000_000_000:
        return f"{value / 1_000_000_000:.1f}Md€"
    elif value >= 1_000_000:
        return f"{value / 1_000_000:.1f}M€"
    elif value >= 1_000:
        return f"{value / 1_000:.0f}k€"
    else:
        return f"{value:.0f}€"


# ---------------------------------------------------------------------------
# FLOWABLES PERSONNALISÉS
# ---------------------------------------------------------------------------

class ColorRect(Flowable):
    """Rectangle coloré (utilisé pour les séparateurs et en-têtes de section)."""
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

def _build_styles():
    """Construit le dictionnaire de styles personnalisés."""
    base = getSampleStyleSheet()
    styles = {}

    styles["title_cover"] = ParagraphStyle(
        "title_cover",
        fontName="Helvetica-Bold",
        fontSize=28,
        textColor=BLANC,
        alignment=TA_LEFT,
        spaceAfter=8,
        leading=34,
    )
    styles["subtitle_cover"] = ParagraphStyle(
        "subtitle_cover",
        fontName="Helvetica",
        fontSize=14,
        textColor=BLEU_CIEL,
        alignment=TA_LEFT,
        spaceAfter=4,
    )
    styles["date_cover"] = ParagraphStyle(
        "date_cover",
        fontName="Helvetica",
        fontSize=10,
        textColor=BLANC,
        alignment=TA_LEFT,
    )
    styles["confidentiel"] = ParagraphStyle(
        "confidentiel",
        fontName="Helvetica-BoldOblique",
        fontSize=9,
        textColor=BLEU_CIEL,
        alignment=TA_LEFT,
    )
    styles["section_title"] = ParagraphStyle(
        "section_title",
        fontName="Helvetica-Bold",
        fontSize=14,
        textColor=BLEU_MARINE,
        spaceBefore=16,
        spaceAfter=8,
    )
    styles["body"] = ParagraphStyle(
        "body",
        fontName="Helvetica",
        fontSize=9,
        textColor=GRIS_TEXTE,
        spaceAfter=4,
        leading=13,
    )
    styles["kpi_label"] = ParagraphStyle(
        "kpi_label",
        fontName="Helvetica",
        fontSize=8,
        textColor=BLANC,
        alignment=TA_CENTER,
    )
    styles["kpi_value"] = ParagraphStyle(
        "kpi_value",
        fontName="Helvetica-Bold",
        fontSize=18,
        textColor=BLEU_CIEL,
        alignment=TA_CENTER,
    )
    styles["table_header"] = ParagraphStyle(
        "table_header",
        fontName="Helvetica-Bold",
        fontSize=8,
        textColor=BLANC,
    )
    styles["table_cell"] = ParagraphStyle(
        "table_cell",
        fontName="Helvetica",
        fontSize=8,
        textColor=BLEU_MARINE,
    )
    styles["footer"] = ParagraphStyle(
        "footer",
        fontName="Helvetica",
        fontSize=7,
        textColor=GRIS_TEXTE,
        alignment=TA_CENTER,
    )
    return styles


# ---------------------------------------------------------------------------
# GRAPHIQUE 1 : PIE CHART — AUM par Type de Client
# ---------------------------------------------------------------------------

def create_chart_aum_by_type(aum_by_type: dict, mode_comex: bool = False) -> io.BytesIO:
    """
    Génère un Donut Chart de la répartition des AUM Funded par type de client.
    Retourne un BytesIO contenant l'image PNG (DPI=150).
    """
    plt.rcParams.update(MPL_PARAMS)

    if not aum_by_type or sum(aum_by_type.values()) == 0:
        # Graphique vide
        fig, ax = plt.subplots(figsize=(5, 5))
        ax.text(0.5, 0.5, "Aucune donnée Funded disponible",
                ha="center", va="center", color=HEX_MARINE, fontsize=11)
        ax.axis("off")
        buf = io.BytesIO()
        fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
        plt.close(fig)
        buf.seek(0)
        return buf

    labels = list(aum_by_type.keys())
    values = list(aum_by_type.values())

    # Palette harmonisée (variations du bleu Amundi)
    palette = [
        HEX_CIEL,
        HEX_MARINE,
        "#1A6B9A",   # bleu intermédiaire
        "#7BC8E8",   # bleu clair pastel
        "#003F7A",   # bleu profond
    ]
    colors = [palette[i % len(palette)] for i in range(len(labels))]

    fig, ax = plt.subplots(figsize=(5.5, 4.5))
    wedges, texts, autotexts = ax.pie(
        values,
        labels=None,
        colors=colors,
        autopct="%1.1f%%",
        startangle=90,
        pctdistance=0.75,
        wedgeprops={"width": 0.55, "edgecolor": "white", "linewidth": 1.5},
    )
    for autotext in autotexts:
        autotext.set_fontsize(9)
        autotext.set_color("white")
        autotext.set_fontweight("bold")

    # Centre du donut : total AUM
    total = sum(values)
    ax.text(0, 0.08, fmt_aum(total), ha="center", va="center",
            fontsize=13, fontweight="bold", color=HEX_MARINE)
    ax.text(0, -0.18, "AUM Funded", ha="center", va="center",
            fontsize=8, color=HEX_GRIS_TEXT)

    # Légende
    legend_patches = [
        mpatches.Patch(color=colors[i], label=f"{labels[i]}: {fmt_aum(values[i])}")
        for i in range(len(labels))
    ]
    ax.legend(
        handles=legend_patches,
        loc="lower center",
        bbox_to_anchor=(0.5, -0.18),
        ncol=2,
        fontsize=8,
        frameon=False,
        labelcolor=HEX_MARINE,
    )
    ax.set_title(
        "Répartition AUM Funded par Type Client",
        fontsize=11, fontweight="bold", color=HEX_MARINE, pad=10
    )

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return buf


# ---------------------------------------------------------------------------
# GRAPHIQUE 2 : BAR CHART HORIZONTAL — Top 10 Deals (Funded AUM)
# ---------------------------------------------------------------------------

def create_chart_top10_deals(top_deals: list, mode_comex: bool = False) -> io.BytesIO:
    """
    Génère un bar chart horizontal des Top 10 deals (funded_aum).
    En mode Comex, les labels utilisent type_client + region.
    Retourne un BytesIO PNG (DPI=150).
    """
    plt.rcParams.update(MPL_PARAMS)

    if not top_deals:
        fig, ax = plt.subplots(figsize=(7, 4))
        ax.text(0.5, 0.5, "Aucun deal Funded disponible",
                ha="center", va="center", color=HEX_MARINE, fontsize=11)
        ax.axis("off")
        buf = io.BytesIO()
        fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
        plt.close(fig)
        buf.seek(0)
        return buf

    # Anonymisation conditionnelle des labels
    labels, values = [], []
    for deal in top_deals[:10]:
        if mode_comex:
            label = f"{deal['type_client']} – {deal['region']}"
        else:
            label = deal["nom_client"]
        labels.append(label)
        values.append(float(deal["funded_aum"]))

    # Tri décroissant
    sorted_pairs = sorted(zip(values, labels), reverse=True)
    values = [v for v, _ in sorted_pairs]
    labels = [l for _, l in sorted_pairs]

    fig, ax = plt.subplots(figsize=(7.5, max(3.5, len(labels) * 0.45 + 1)))

    y_pos = range(len(labels))
    bars = ax.barh(
        y_pos, values,
        color=HEX_CIEL,
        edgecolor="white",
        linewidth=0.5,
        height=0.6,
    )

    # Annotations valeurs
    for bar, val in zip(bars, values):
        ax.text(
            bar.get_width() + max(values) * 0.01,
            bar.get_y() + bar.get_height() / 2,
            fmt_aum(val),
            va="center", ha="left",
            fontsize=8, color=HEX_MARINE, fontweight="bold"
        )

    ax.set_yticks(y_pos)
    ax.set_yticklabels(labels, fontsize=8.5, color=HEX_MARINE)
    ax.invert_yaxis()
    ax.set_xlabel("AUM Funded (€)", fontsize=9, color=HEX_MARINE)
    ax.set_title(
        "Top 10 Deals — AUM Financé",
        fontsize=11, fontweight="bold", color=HEX_MARINE, pad=10
    )

    # Formatage de l'axe X
    ax.xaxis.set_major_formatter(
        plt.FuncFormatter(lambda x, _: fmt_aum(x))
    )
    ax.set_xlim(0, max(values) * 1.18)
    ax.tick_params(axis="x", labelsize=8, colors=HEX_MARINE)
    ax.grid(axis="x", alpha=0.4)

    # Barre de fond légère
    for i, bar in enumerate(bars):
        ax.barh(
            bar.get_y() + bar.get_height() / 2,
            max(values) * 1.15,
            height=bar.get_height(),
            color=HEX_GRIS if i % 2 == 0 else "white",
            alpha=0.3,
            zorder=0,
        )

    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return buf


# ---------------------------------------------------------------------------
# GRAPHIQUE 3 : BAR CHART — AUM Target vs Funded par Statut actif
# ---------------------------------------------------------------------------

def create_chart_target_vs_funded(pipeline_df: pd.DataFrame, mode_comex: bool = False) -> io.BytesIO:
    """
    Graphique de comparaison Target AUM Initial vs Funded AUM par deal actif.
    """
    plt.rcParams.update(MPL_PARAMS)

    # Filtrer les deals significatifs (Funded ou Soft Commit)
    df = pipeline_df[pipeline_df["statut"].isin(["Funded", "Soft Commit", "Due Diligence"])].copy()
    df = df[df["target_aum_initial"] > 0].head(8)

    if df.empty:
        fig, ax = plt.subplots(figsize=(7, 3))
        ax.text(0.5, 0.5, "Données insuffisantes", ha="center", va="center",
                color=HEX_MARINE)
        ax.axis("off")
        buf = io.BytesIO()
        fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
        plt.close(fig)
        buf.seek(0)
        return buf

    if mode_comex:
        df["label"] = df.apply(lambda r: f"{r['type_client']} – {r['region']}", axis=1)
    else:
        df["label"] = df["nom_client"].str[:20]

    x = np.arange(len(df))
    width = 0.35

    fig, ax = plt.subplots(figsize=(8, 4))

    ax.bar(x - width/2, df["target_aum_initial"], width, label="Target Initial",
           color=HEX_GRIS, edgecolor="white", linewidth=0.5)
    ax.bar(x + width/2, df["funded_aum"], width, label="Funded AUM",
           color=HEX_CIEL, edgecolor="white", linewidth=0.5)

    ax.set_xticks(x)
    ax.set_xticklabels(df["label"], rotation=35, ha="right", fontsize=7.5, color=HEX_MARINE)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: fmt_aum(y)))
    ax.tick_params(axis="y", labelsize=8, colors=HEX_MARINE)
    ax.legend(fontsize=8, frameon=False, labelcolor=HEX_MARINE)
    ax.set_title("Target Initial vs AUM Financé", fontsize=10,
                 fontweight="bold", color=HEX_MARINE, pad=8)
    ax.set_ylabel("AUM (€)", fontsize=9, color=HEX_MARINE)
    ax.grid(axis="y", alpha=0.3)

    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return buf


# ---------------------------------------------------------------------------
# PAGE DE GARDE
# ---------------------------------------------------------------------------

def _build_cover_page(styles: dict, mode_comex: bool, kpis: dict) -> list:
    """Construit les éléments de la page de garde."""
    elements = []
    today_str = date.today().strftime("%d %B %Y")

    # Bandeau bleu marine pleine largeur (simulé avec une table)
    cover_data = [[
        Paragraph(
            f"""
            <font color='#00A8E1' size='9'>CONFIDENTIEL{'  |  MODE COMEX ACTIVÉ' if mode_comex else ''}</font><br/>
            <br/>
            <font color='white' size='26'><b>Executive Report</b></font><br/>
            <font color='white' size='20'>Pipeline & Performance</font><br/>
            <br/>
            <font color='#00A8E1' size='11'>Asset Management Division</font><br/>
            <br/>
            <font color='white' size='9'>{today_str}</font>
            """,
            ParagraphStyle("cover_inner", fontName="Helvetica", fontSize=9,
                           textColor=BLANC, leading=16, spaceBefore=0)
        )
    ]]

    cover_table = Table(cover_data, colWidths=[PAGE_W - 4*cm])
    cover_table.setStyle(TableStyle([
        ("BACKGROUND",  (0, 0), (-1, -1), BLEU_MARINE),
        ("TOPPADDING",  (0, 0), (-1, -1), 40),
        ("BOTTOMPADDING",(0, 0),(-1, -1), 50),
        ("LEFTPADDING", (0, 0), (-1, -1), 30),
        ("RIGHTPADDING",(0, 0), (-1, -1), 20),
        ("VALIGN",      (0, 0), (-1, -1), "MIDDLE"),
    ]))
    elements.append(cover_table)
    elements.append(Spacer(1, 0.5*cm))

    # Bandeau bleu ciel accent
    elements.append(ColorRect(PAGE_W - 4*cm, 4, BLEU_CIEL))
    elements.append(Spacer(1, 0.7*cm))

    # Résumé des KPIs sur la cover
    kpi_items = [
        ("AUM Financé Total",  fmt_aum(kpis.get("total_funded", 0))),
        ("Pipeline Actif",     fmt_aum(kpis.get("pipeline_actif", 0))),
        ("Taux Conversion",    f"{kpis.get('taux_conversion', 0):.1f}%"),
        ("Deals Actifs",       str(kpis.get("nb_deals_actifs", 0))),
    ]

    kpi_row_labels = [[Paragraph(item[0], ParagraphStyle(
        "kl", fontName="Helvetica", fontSize=8, textColor=BLANC, alignment=TA_CENTER
    )) for item in kpi_items]]
    kpi_row_values = [[Paragraph(item[1], ParagraphStyle(
        "kv", fontName="Helvetica-Bold", fontSize=17, textColor=BLEU_CIEL, alignment=TA_CENTER
    )) for item in kpi_items]]

    kpi_data = kpi_row_labels + kpi_row_values
    col_w = (PAGE_W - 4*cm) / 4
    kpi_table = Table(kpi_data, colWidths=[col_w] * 4, rowHeights=[1.0*cm, 1.2*cm])
    kpi_table.setStyle(TableStyle([
        ("BACKGROUND",   (0, 0), (-1, -1), BLEU_MARINE),
        ("TOPPADDING",   (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 8),
        ("LEFTPADDING",  (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("LINEAFTER",    (0, 0), (2, -1),  0.5, BLEU_CIEL),
        ("VALIGN",       (0, 0), (-1, -1), "MIDDLE"),
    ]))
    elements.append(kpi_table)
    elements.append(Spacer(1, 0.4*cm))

    disclaimer = (
        "Ce document est strictement confidentiel et destiné exclusivement à l'usage interne. "
        "Toute reproduction ou diffusion externe est strictement interdite. "
        "Les performances passées ne préjugent pas des performances futures."
    )
    if mode_comex:
        disclaimer += " — MODE COMEX : Les noms des clients ont été anonymisés."

    elements.append(Paragraph(disclaimer, ParagraphStyle(
        "disc", fontName="Helvetica-Oblique", fontSize=7, textColor=GRIS_TEXTE,
        alignment=TA_CENTER, leading=10
    )))

    elements.append(PageBreak())
    return elements


# ---------------------------------------------------------------------------
# SECTION : TABLEAU PIPELINE ACTIF
# ---------------------------------------------------------------------------

def _build_pipeline_table(pipeline_df: pd.DataFrame, styles: dict, mode_comex: bool) -> list:
    """Construit le tableau récapitulatif du pipeline actif."""
    elements = []

    elements.append(Paragraph("Pipeline Actif — Récapitulatif", styles["section_title"]))
    elements.append(ColorRect(PAGE_W - 4*cm, 2, BLEU_CIEL))
    elements.append(Spacer(1, 0.3*cm))

    # Filtrer les statuts actifs
    active_statuts = ["Prospect", "Initial Pitch", "Due Diligence", "Soft Commit"]
    df = pipeline_df[pipeline_df["statut"].isin(active_statuts)].copy()

    if df.empty:
        elements.append(Paragraph("Aucun deal actif dans le pipeline.", styles["body"]))
        return elements

    # Anonymisation si Mode Comex
    if mode_comex:
        df["nom_client"] = df.apply(lambda r: f"{r['type_client']} – {r['region']}", axis=1)

    # En-tête
    headers = ["Client", "Fonds", "Statut", "AUM Cible", "AUM Révisé", "Prochaine Action"]
    header_cells = [Paragraph(h, ParagraphStyle(
        "th", fontName="Helvetica-Bold", fontSize=7.5, textColor=BLANC
    )) for h in headers]
    table_data = [header_cells]

    # Lignes de données
    today = date.today()
    for _, row in df.iterrows():
        # Alerte si next_action_date dépassée
        next_action = str(row.get("next_action_date", "") or "")
        try:
            nad = date.fromisoformat(next_action)
            if nad < today:
                next_action_str = f"⚠ {next_action}"
            else:
                next_action_str = next_action
        except Exception:
            next_action_str = next_action or "—"

        cell_style = ParagraphStyle(
            "td", fontName="Helvetica", fontSize=7.5, textColor=BLEU_MARINE
        )
        table_data.append([
            Paragraph(str(row.get("nom_client", ""))[:28], cell_style),
            Paragraph(str(row.get("fonds", "")), cell_style),
            Paragraph(str(row.get("statut", "")), cell_style),
            Paragraph(fmt_aum(float(row.get("target_aum_initial", 0) or 0)), cell_style),
            Paragraph(fmt_aum(float(row.get("revised_aum", 0) or 0)), cell_style),
            Paragraph(next_action_str, cell_style),
        ])

    col_widths = [4.2*cm, 3.0*cm, 2.2*cm, 2.2*cm, 2.2*cm, 2.7*cm]
    pipeline_table = Table(table_data, colWidths=col_widths, repeatRows=1)

    # Style de la table
    table_style = TableStyle([
        # En-tête
        ("BACKGROUND",    (0, 0), (-1, 0),  BLEU_MARINE),
        ("TEXTCOLOR",     (0, 0), (-1, 0),  BLANC),
        ("ALIGN",         (0, 0), (-1, -1), "LEFT"),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",    (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING",   (0, 0), (-1, -1), 6),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 6),
        ("GRID",          (0, 0), (-1, -1), 0.4, GRIS_CLAIR),
        ("LINEBELOW",     (0, 0), (-1, 0),  1.5, BLEU_CIEL),
    ])
    # Alternance des lignes
    for i in range(1, len(table_data)):
        if i % 2 == 0:
            table_style.add("BACKGROUND", (0, i), (-1, i), HexColor("#F5F8FC"))

    pipeline_table.setStyle(table_style)
    elements.append(pipeline_table)
    return elements


# ---------------------------------------------------------------------------
# CALLBACK PIED DE PAGE / EN-TÊTE
# ---------------------------------------------------------------------------

def _header_footer(canvas, doc):
    """Dessine l'en-tête et le pied de page sur chaque page (sauf cover)."""
    canvas.saveState()
    page_num = doc.page

    if page_num > 1:
        # Barre bleu marine en bas
        canvas.setFillColor(BLEU_MARINE)
        canvas.rect(0, 0, PAGE_W, 1.1*cm, fill=1, stroke=0)

        # Pied de page texte
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(GRIS_CLAIR)
        canvas.drawString(2*cm, 0.38*cm, "Executive Report — Asset Management Division")
        canvas.drawRightString(PAGE_W - 2*cm, 0.38*cm, f"Page {page_num}")

        # Filet bleu ciel en haut
        canvas.setStrokeColor(BLEU_CIEL)
        canvas.setLineWidth(3)
        canvas.line(0, PAGE_H - 0.5*cm, PAGE_W, PAGE_H - 0.5*cm)

        # Titre en haut à droite
        canvas.setFont("Helvetica", 7.5)
        canvas.setFillColor(BLEU_MARINE)
        canvas.drawRightString(PAGE_W - 2*cm, PAGE_H - 0.35*cm, "CONFIDENTIEL")

    canvas.restoreState()


# ---------------------------------------------------------------------------
# FONCTION PRINCIPALE : generate_pdf
# ---------------------------------------------------------------------------

def generate_pdf(
    pipeline_df: pd.DataFrame,
    kpis: dict,
    mode_comex: bool = False,
) -> bytes:
    """
    Génère le rapport PDF complet style Pitchbook.

    Args:
        pipeline_df: DataFrame du pipeline (enrichi clients).
        kpis: Dict des indicateurs clés (depuis database.get_kpis()).
        mode_comex: Si True, anonymise TOUS les noms avant la génération.

    Returns:
        bytes du fichier PDF prêt à télécharger.
    """
    # =========================================================
    # ÉTAPE 0 : Anonymisation COMPLÈTE si Mode Comex
    # (aucun nom ne doit traverser les fonctions suivantes)
    # =========================================================
    if mode_comex:
        pipeline_df = anonymize_df(pipeline_df)
        # Anonymiser aussi les top_deals dans kpis
        top_deals_safe = []
        for deal in kpis.get("top_deals", []):
            d = dict(deal)
            d["nom_client"] = f"{d['type_client']} – {d['region']}"
            top_deals_safe.append(d)
        kpis = {**kpis, "top_deals": top_deals_safe}
    else:
        top_deals_safe = kpis.get("top_deals", [])

    styles = _build_styles()

    # Buffer mémoire pour le PDF
    pdf_buffer = io.BytesIO()

    doc = SimpleDocTemplate(
        pdf_buffer,
        pagesize=A4,
        leftMargin=2*cm,
        rightMargin=2*cm,
        topMargin=1.8*cm,
        bottomMargin=1.8*cm,
        title="Executive Report — Asset Management",
        author="Asset Management Division",
    )

    elements = []

    # =========================================================
    # PAGE 1 : PAGE DE GARDE
    # =========================================================
    elements += _build_cover_page(styles, mode_comex, kpis)

    # =========================================================
    # PAGE 2 : GRAPHIQUES — AUM par Type + Top 10 Deals
    # =========================================================
    elements.append(Paragraph("Analyse du Pipeline", styles["section_title"]))
    elements.append(ColorRect(PAGE_W - 4*cm, 2, BLEU_CIEL))
    elements.append(Spacer(1, 0.4*cm))

    # Génération des graphiques Matplotlib → BytesIO → ReportLab Image
    chart_width  = (PAGE_W - 5*cm) / 2   # 2 colonnes
    chart_height = chart_width * 0.85

    buf_pie  = create_chart_aum_by_type(kpis.get("aum_by_type", {}), mode_comex)
    buf_bar  = create_chart_top10_deals(top_deals_safe, mode_comex)

    img_pie = Image(buf_pie, width=chart_width, height=chart_height)
    img_bar = Image(buf_bar, width=chart_width, height=chart_height)

    # Disposition côte à côte
    charts_table = Table(
        [[img_pie, img_bar]],
        colWidths=[chart_width + 0.2*cm, chart_width + 0.2*cm]
    )
    charts_table.setStyle(TableStyle([
        ("ALIGN",   (0, 0), (-1, -1), "CENTER"),
        ("VALIGN",  (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING",  (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ]))
    elements.append(charts_table)
    elements.append(Spacer(1, 0.5*cm))

    # Graphique Target vs Funded (pleine largeur)
    elements.append(Paragraph("Target vs AUM Financé (deals actifs)", styles["section_title"]))
    elements.append(ColorRect(PAGE_W - 4*cm, 2, BLEU_CIEL))
    elements.append(Spacer(1, 0.3*cm))

    buf_tvf = create_chart_target_vs_funded(pipeline_df, mode_comex)
    img_tvf = Image(buf_tvf, width=PAGE_W - 5*cm, height=(PAGE_W - 5*cm) * 0.42)
    elements.append(img_tvf)
    elements.append(Spacer(1, 0.5*cm))

    # =========================================================
    # PAGE 3 : TABLEAU PIPELINE ACTIF
    # =========================================================
    elements.append(PageBreak())
    elements += _build_pipeline_table(pipeline_df, styles, mode_comex)
    elements.append(Spacer(1, 0.6*cm))

    # Récapitulatif des deals Lost/Paused
    lost_paused = pipeline_df[pipeline_df["statut"].isin(["Lost", "Paused"])].copy()
    if not lost_paused.empty:
        elements.append(Paragraph("Deals Perdus / En Pause", styles["section_title"]))
        elements.append(ColorRect(PAGE_W - 4*cm, 2, GRIS_CLAIR))
        elements.append(Spacer(1, 0.3*cm))

        if mode_comex:
            lost_paused["nom_client"] = lost_paused.apply(
                lambda r: f"{r['type_client']} – {r['region']}", axis=1
            )

        lp_headers = ["Client", "Fonds", "Statut", "Raison", "Concurrent"]
        lp_data = [[Paragraph(h, ParagraphStyle(
            "lph", fontName="Helvetica-Bold", fontSize=7.5, textColor=BLANC
        )) for h in lp_headers]]

        for _, row in lost_paused.iterrows():
            lp_data.append([
                Paragraph(str(row.get("nom_client", ""))[:25],
                          ParagraphStyle("lptd", fontName="Helvetica", fontSize=7.5, textColor=GRIS_TEXTE)),
                Paragraph(str(row.get("fonds", "")),
                          ParagraphStyle("lptd2", fontName="Helvetica", fontSize=7.5, textColor=GRIS_TEXTE)),
                Paragraph(str(row.get("statut", "")),
                          ParagraphStyle("lptd3", fontName="Helvetica-Bold", fontSize=7.5, textColor=HexColor("#888888"))),
                Paragraph(str(row.get("raison_perte", "") or "—"),
                          ParagraphStyle("lptd4", fontName="Helvetica", fontSize=7.5, textColor=GRIS_TEXTE)),
                Paragraph(str(row.get("concurrent_choisi", "") or "—"),
                          ParagraphStyle("lptd5", fontName="Helvetica", fontSize=7.5, textColor=GRIS_TEXTE)),
            ])

        lp_table = Table(
            lp_data,
            colWidths=[4.0*cm, 3.0*cm, 2.0*cm, 3.0*cm, 3.5*cm],
            repeatRows=1
        )
        lp_table.setStyle(TableStyle([
            ("BACKGROUND",   (0, 0), (-1, 0),  HexColor("#888888")),
            ("TEXTCOLOR",    (0, 0), (-1, 0),  BLANC),
            ("GRID",         (0, 0), (-1, -1), 0.4, GRIS_CLAIR),
            ("TOPPADDING",   (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING",(0, 0), (-1, -1), 5),
            ("LEFTPADDING",  (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ("LINEBELOW",    (0, 0), (-1, 0),  1.0, GRIS_CLAIR),
        ]))
        elements.append(lp_table)

    # =========================================================
    # CONSTRUCTION DU PDF
    # =========================================================
    doc.build(elements, onFirstPage=_header_footer, onLaterPages=_header_footer)
    pdf_bytes = pdf_buffer.getvalue()
    pdf_buffer.close()

    return pdf_bytes
