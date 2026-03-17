# =============================================================================
# app.py — CRM Asset Management — Amundi Edition  v15.0 — Universal Export Hub
# Architecture : Full Modal (@st.dialog) + Split-Screen CRM
# Session-state keys are namespaced per feature to prevent cross-tab conflicts
# =============================================================================

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from datetime import date, timedelta
import io
import sys
import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import database as db
import pdf_generator as pdf_gen

# ---------------------------------------------------------------------------
# CONSTANTES
# ---------------------------------------------------------------------------
MARINE  = "#001c4b"
CIEL    = "#019ee1"
ORANGE  = "#f07d00"
BLANC   = "#ffffff"
GRIS    = "#e8e8e8"
GTXT    = "#444444"
B_MID   = "#1a5e8a"
B_PAL   = "#4a8fbd"
B_DEP   = "#003f7a"
PALETTE = [B_MID, MARINE, B_PAL, B_DEP, "#2c7fb8",
           "#004f8c", "#6baed6", "#08519c", "#9ecae1", "#003060"]

TYPES_CLIENT      = ["IFA", "Wholesale", "Instit", "Family Office"]
REGIONS           = db.REGIONS_REFERENTIEL
FONDS             = db.FONDS_REFERENTIEL
STATUTS           = ["Prospect","Initial Pitch","Due Diligence",
                     "Soft Commit","Funded","Paused","Lost","Redeemed"]
STATUTS_ACTIFS    = ["Prospect","Initial Pitch","Due Diligence","Soft Commit"]
RAISONS_PERTE     = ["Pricing","Track Record","Macro","Competitor","Autre"]
TYPES_INTERACTION = ["Call","Meeting","Email","Roadshow","Conference","Autre"]

COUNTRIES_LIST = [
    "", "United Arab Emirates", "Saudi Arabia", "Qatar", "Kuwait", "Bahrain", "Oman",
    "United Kingdom", "France", "Germany", "Switzerland", "Luxembourg", "Netherlands",
    "Italy", "Spain", "Belgium", "Austria", "Sweden", "Norway", "Denmark", "Finland",
    "Singapore", "Japan", "Hong Kong", "China", "South Korea", "Australia", "India",
    "United States", "Canada", "Brazil", "Mexico", "South Africa", "Egypt",
]

STATUT_COLORS = {
    "Funded":        B_MID,    "Soft Commit":   "#2c7fb8",
    "Due Diligence": "#004f8c", "Initial Pitch": B_PAL,
    "Prospect":      "#9ecae1", "Lost":          "#aaaaaa",
    "Paused":        "#c0c0c0", "Redeemed":      "#b8b8d0",
}

fmt_m = db.format_finance

# ---------------------------------------------------------------------------
# PPTX — Account Review Generator (v14.0)
# ---------------------------------------------------------------------------
def _rgb(hex_str):
    """Convert #RRGGBB to RGBColor."""
    h = hex_str.lstrip("#")
    return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

def generate_account_review_pptx(client_data: dict, group_summary: dict,
                                   contacts_df, activities_df) -> bytes:
    """
    Generates a 2-slide PPTX Account Review.
    Slide 1: Title — Revue de Compte Stratégique
    Slide 2: Synthèse — AUM, KYC, Fonds, Next Actions
    Returns bytes ready for st.download_button.
    """
    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)

    MARINE_RGB  = _rgb("#001c4b")
    CIEL_RGB    = _rgb("#019ee1")
    ORANGE_RGB  = _rgb("#f07d00")
    BLANC_RGB   = _rgb("#ffffff")
    LIGHT_RGB   = _rgb("#f2f6fa")
    GREY_RGB    = _rgb("#666666")

    blank_layout = prs.slide_layouts[6]  # Completely blank

    def _add_rect(slide, x, y, w, h, fill_rgb, line_rgb=None, line_width_pt=0):
        from pptx.util import Pt as _Pt
        shape = slide.shapes.add_shape(
            1,  # MSO_SHAPE_TYPE.RECTANGLE = 1
            Inches(x), Inches(y), Inches(w), Inches(h))
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_rgb
        if line_rgb:
            shape.line.color.rgb = line_rgb
            shape.line.width = Pt(line_width_pt)
        else:
            shape.line.fill.background()
        return shape

    def _add_text(slide, text, x, y, w, h, font_size, bold=False, color_rgb=None,
                  align=PP_ALIGN.LEFT, italic=False, wrap=True):
        txBox = slide.shapes.add_textbox(
            Inches(x), Inches(y), Inches(w), Inches(h))
        txBox.word_wrap = wrap
        tf = txBox.text_frame
        tf.word_wrap = wrap
        para = tf.paragraphs[0]
        para.alignment = align
        run = para.add_run()
        run.text = str(text)
        run.font.size = Pt(font_size)
        run.font.bold = bold
        run.font.italic = italic
        if color_rgb:
            run.font.color.rgb = color_rgb
        return txBox

    # ── SLIDE 1 — TITLE ──────────────────────────────────────────────────────
    slide1 = prs.slides.add_slide(blank_layout)

    # Dark header bar
    _add_rect(slide1, 0, 0, 13.33, 2.8, MARINE_RGB)
    # Accent stripe
    _add_rect(slide1, 0, 2.8, 13.33, 0.08, CIEL_RGB)

    # Company name
    _add_text(slide1,
              "REVUE DE COMPTE STRATÉGIQUE",
              0.6, 0.35, 12.0, 0.7,
              font_size=13, bold=False, color_rgb=CIEL_RGB,
              align=PP_ALIGN.LEFT)
    _add_text(slide1,
              client_data.get("nom_client", "—"),
              0.6, 0.95, 12.0, 1.2,
              font_size=40, bold=True, color_rgb=BLANC_RGB,
              align=PP_ALIGN.LEFT)

    # Date
    _add_text(slide1,
              date.today().strftime("%d %B %Y").upper(),
              0.6, 2.2, 6.0, 0.5,
              font_size=11, bold=False, color_rgb=_rgb("#7ab8d8"),
              align=PP_ALIGN.LEFT)

    # KYC badge (bottom-right of header)
    kyc = client_data.get("kyc_status", "En cours")
    kyc_colors = {"Validé": _rgb("#22a062"), "En cours": ORANGE_RGB, "Bloqué": _rgb("#c0392b")}
    kyc_rgb = kyc_colors.get(kyc, GREY_RGB)
    _add_rect(slide1, 10.0, 1.8, 2.7, 0.7, kyc_rgb)
    _add_text(slide1, "KYC : {}".format(kyc),
              10.0, 1.9, 2.7, 0.5,
              font_size=14, bold=True, color_rgb=BLANC_RGB,
              align=PP_ALIGN.CENTER)

    # Info grid below header
    infos = [
        ("TYPE", client_data.get("type_client","—")),
        ("RÉGION", client_data.get("region","—")),
        ("TIER", client_data.get("tier","—")),
        ("COUNTRY", client_data.get("country","—") or "—"),
    ]
    for idx, (lbl, val) in enumerate(infos):
        xi = 0.5 + idx * 3.2
        _add_rect(slide1, xi, 3.1, 3.0, 1.2, LIGHT_RGB)
        _add_text(slide1, lbl, xi+0.15, 3.2, 2.7, 0.4,
                  font_size=9, bold=False, color_rgb=GREY_RGB)
        _add_text(slide1, val, xi+0.15, 3.55, 2.7, 0.6,
                  font_size=16, bold=True, color_rgb=MARINE_RGB)

    # Product interests
    prods = [p.strip() for p in str(client_data.get("product_interests","")).split(",") if p.strip()]
    if prods:
        _add_text(slide1, "Product Interests : " + "  ·  ".join(prods),
                  0.5, 4.5, 12.3, 0.5,
                  font_size=11, bold=False, color_rgb=GREY_RGB)

    # Footer
    _add_rect(slide1, 0, 7.1, 13.33, 0.4, _rgb("#f0f4f8"))
    _add_text(slide1, "Document à usage interne — Confidentiel — Amundi Asset Management",
              0.5, 7.15, 12.0, 0.3,
              font_size=8, bold=False, color_rgb=GREY_RGB, align=PP_ALIGN.CENTER)

    # ── SLIDE 2 — SYNTHÈSE ───────────────────────────────────────────────────
    slide2 = prs.slides.add_slide(blank_layout)

    # Thin top bar
    _add_rect(slide2, 0, 0, 13.33, 0.55, MARINE_RGB)
    _add_text(slide2,
              "SYNTHÈSE  —  {}".format(client_data.get("nom_client","").upper()),
              0.4, 0.1, 10.0, 0.38,
              font_size=11, bold=True, color_rgb=BLANC_RGB)
    _add_text(slide2, date.today().strftime("%d/%m/%Y"),
              11.0, 0.1, 2.0, 0.38,
              font_size=11, bold=False, color_rgb=_rgb("#7ab8d8"),
              align=PP_ALIGN.RIGHT)

    # ── KPI cards row ────────────────────────────────────────────────────────
    kpi_data = [
        ("AUM Consolidé",  db.format_finance(group_summary["aum_consolide"]), MARINE_RGB),
        ("AUM Direct",     db.format_finance(group_summary["aum_direct"]),    CIEL_RGB),
        ("Fonds Investis", str(len(group_summary["fonds_investis"])),          _rgb("#1a5e8a")),
        ("Deals Actifs",   str(len(group_summary["next_actions"])),            ORANGE_RGB),
    ]
    for idx, (lbl, val, clr) in enumerate(kpi_data):
        xi = 0.3 + idx * 3.2
        _add_rect(slide2, xi, 0.75, 3.0, 1.3, MARINE_RGB)
        _add_rect(slide2, xi, 0.75, 3.0, 0.06, clr)  # color accent top
        _add_text(slide2, lbl, xi+0.15, 0.85, 2.7, 0.4,
                  font_size=9, bold=False, color_rgb=_rgb("#7ab8d8"))
        _add_text(slide2, val, xi+0.15, 1.2, 2.7, 0.65,
                  font_size=22, bold=True, color_rgb=BLANC_RGB)

    # ── Fonds Investis ───────────────────────────────────────────────────────
    _add_text(slide2, "FONDS INVESTIS", 0.3, 2.25, 5.5, 0.35,
              font_size=9, bold=True, color_rgb=CIEL_RGB)
    _add_rect(slide2, 0.3, 2.6, 5.8, 0.04, CIEL_RGB)
    if group_summary["fonds_investis"]:
        for fi, fonds_nm in enumerate(group_summary["fonds_investis"]):
            _add_rect(slide2, 0.3, 2.72 + fi*0.42, 5.8, 0.38, LIGHT_RGB)
            _add_text(slide2, fonds_nm, 0.5, 2.74 + fi*0.42, 5.4, 0.36,
                      font_size=12, bold=False, color_rgb=MARINE_RGB)
    else:
        _add_text(slide2, "Aucun fonds financé", 0.3, 2.72, 5.8, 0.4,
                  font_size=11, bold=False, color_rgb=GREY_RGB)

    # ── Prochaines Actions ───────────────────────────────────────────────────
    _add_text(slide2, "PROCHAINES ACTIONS", 6.8, 2.25, 6.2, 0.35,
              font_size=9, bold=True, color_rgb=ORANGE_RGB)
    _add_rect(slide2, 6.8, 2.6, 6.2, 0.04, ORANGE_RGB)
    if group_summary["next_actions"]:
        for ai, act in enumerate(group_summary["next_actions"][:4]):
            row_y = 2.72 + ai * 0.55
            _add_rect(slide2, 6.8, row_y, 6.2, 0.5, LIGHT_RGB)
            label = "{} — {}  |  {}  |  {}".format(
                act["fonds"], act["statut"],
                db.format_finance(act["aum_pipeline"]),
                act["nad"])
            _add_text(slide2, label, 7.0, row_y + 0.07, 5.8, 0.4,
                      font_size=10, bold=False, color_rgb=MARINE_RGB)
    else:
        _add_text(slide2, "Aucune action planifiée", 6.8, 2.72, 6.2, 0.4,
                  font_size=11, bold=False, color_rgb=GREY_RGB)

    # ── Dernières Activités ──────────────────────────────────────────────────
    _add_text(slide2, "DERNIÈRES ACTIVITÉS", 0.3, 5.3, 6.5, 0.35,
              font_size=9, bold=True, color_rgb=MARINE_RGB)
    _add_rect(slide2, 0.3, 5.65, 12.7, 0.03, _rgb("#e8e8e8"))
    TYPE_COLORS_PPTX = {
        "Call": CIEL_RGB, "Meeting": _rgb("#1a5e8a"),
        "Email": _rgb("#4a8fbd"), "Autre": GREY_RGB,
    }
    if not activities_df.empty:
        for ri, (_, act_row) in enumerate(activities_df.head(2).iterrows()):
            row_y = 5.75 + ri * 0.5
            a_typ = str(act_row.get("type_interaction",""))
            a_col = TYPE_COLORS_PPTX.get(a_typ, GREY_RGB)
            _add_rect(slide2, 0.3, row_y, 0.06, 0.38, a_col)
            _add_text(slide2,
                      "[{}]  {}  —  {}".format(
                          a_typ,
                          str(act_row.get("date","")),
                          str(act_row.get("notes",""))[:80]),
                      0.5, row_y + 0.02, 12.3, 0.36,
                      font_size=10, bold=False, color_rgb=MARINE_RGB)
    else:
        _add_text(slide2, "Aucune activité enregistrée.", 0.3, 5.75, 12.7, 0.4,
                  font_size=11, bold=False, color_rgb=GREY_RGB)

    # Footer
    _add_rect(slide2, 0, 7.1, 13.33, 0.4, _rgb("#f0f4f8"))
    _add_text(slide2, "Document à usage interne — Confidentiel — Amundi Asset Management",
              0.5, 7.15, 12.0, 0.3,
              font_size=8, bold=False, color_rgb=GREY_RGB, align=PP_ALIGN.CENTER)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# PPTX — Global Portfolio Report Generator (v15.0)
# ---------------------------------------------------------------------------
def generate_global_pptx(kpis: dict, pipeline_df, mode_comex: bool = False) -> bytes:
    """
    2-slide global PPTX report:
      Slide 1 : Title & KPI Overview
      Slide 2 : Top Deals table (client names masked if mode_comex)
    Returns bytes ready for st.download_button.
    """
    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)

    MARINE_RGB  = _rgb("#001c4b")
    CIEL_RGB    = _rgb("#019ee1")
    ORANGE_RGB  = _rgb("#f07d00")
    BLANC_RGB   = _rgb("#ffffff")
    LIGHT_RGB   = _rgb("#f2f6fa")
    GREY_RGB    = _rgb("#666666")

    blank_layout = prs.slide_layouts[6]

    def _rect(slide, x, y, w, h, fill_rgb):
        sh = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
        sh.fill.solid(); sh.fill.fore_color.rgb = fill_rgb
        sh.line.fill.background()
        return sh

    def _text(slide, text, x, y, w, h, fs, bold=False, color=None, align=PP_ALIGN.LEFT):
        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]; p.alignment = align
        r = p.add_run(); r.text = str(text)
        r.font.size = Pt(fs); r.font.bold = bold
        if color: r.font.color.rgb = color
        return tb

    # ── SLIDE 1 — TITRE & KPIs ───────────────────────────────────────────────
    s1 = prs.slides.add_slide(blank_layout)
    _rect(s1, 0, 0, 13.33, 2.4, MARINE_RGB)
    _rect(s1, 0, 2.4, 13.33, 0.07, CIEL_RGB)

    label_comex = " [MODE COMEX — ANONYMISÉ]" if mode_comex else ""
    _text(s1, "RAPPORT GLOBAL — PORTFOLIO PIPELINE" + label_comex,
          0.6, 0.3, 12.0, 0.6, 11, color=CIEL_RGB)
    _text(s1, "Amundi Asset Management",
          0.6, 0.85, 12.0, 0.9, 34, bold=True, color=BLANC_RGB)
    _text(s1, date.today().strftime("%d %B %Y").upper(),
          0.6, 1.8, 6.0, 0.5, 11, color=_rgb("#7ab8d8"))

    # KPI row
    kpi_items = [
        ("AUM Financé Total",  fmt_m(kpis.get("total_funded", 0)),    CIEL_RGB),
        ("Pipeline Actif",     fmt_m(kpis.get("pipeline_actif", 0)),   _rgb("#1a5e8a")),
        ("Pipeline Pondéré",   fmt_m(kpis.get("weighted_pipeline",0)), ORANGE_RGB),
        ("Taux Conversion",    "{:.1f}%".format(kpis.get("taux_conversion",0)), _rgb("#22a062")),
    ]
    for idx, (lbl, val, clr) in enumerate(kpi_items):
        xi = 0.4 + idx * 3.2
        _rect(s1, xi, 2.65, 3.0, 1.4, MARINE_RGB)
        _rect(s1, xi, 2.65, 3.0, 0.07, clr)
        _text(s1, lbl, xi+0.15, 2.78, 2.7, 0.4, 9, color=_rgb("#7ab8d8"))
        _text(s1, val, xi+0.15, 3.1,  2.7, 0.75, 22, bold=True, color=BLANC_RGB)

    # Statut pills
    statut_rep = kpis.get("statut_repartition", {})
    statut_items = [(s, statut_rep.get(s, 0)) for s in
                    ["Funded","Soft Commit","Due Diligence","Initial Pitch","Prospect"]
                    if statut_rep.get(s, 0) > 0]
    _text(s1, "RÉPARTITION PAR STATUT", 0.4, 4.25, 12.5, 0.35, 8, bold=True, color=GREY_RGB)
    for si, (stat, cnt) in enumerate(statut_items):
        xi = 0.4 + si * 2.55
        _rect(s1, xi, 4.65, 2.35, 0.9, LIGHT_RGB)
        _text(s1, stat[:12], xi+0.1, 4.72, 2.15, 0.35, 9, color=GREY_RGB)
        _text(s1, str(cnt), xi+0.1, 5.0, 2.15, 0.45, 18, bold=True, color=MARINE_RGB)

    # Footer
    _rect(s1, 0, 7.1, 13.33, 0.4, LIGHT_RGB)
    _text(s1, "Document à usage interne — Confidentiel — Amundi Asset Management",
          0.5, 7.15, 12.0, 0.3, 8, color=GREY_RGB, align=PP_ALIGN.CENTER)

    # ── SLIDE 2 — TOP DEALS ──────────────────────────────────────────────────
    s2 = prs.slides.add_slide(blank_layout)
    _rect(s2, 0, 0, 13.33, 0.55, MARINE_RGB)
    _text(s2, "TOP DEALS — AUM FINANCÉ", 0.4, 0.1, 10.0, 0.38, 11, bold=True, color=BLANC_RGB)
    _text(s2, date.today().strftime("%d/%m/%Y"), 11.0, 0.1, 2.0, 0.38,
          11, color=_rgb("#7ab8d8"), align=PP_ALIGN.RIGHT)

    if mode_comex:
        _text(s2, "Noms de clients masqués — Mode Comex",
              0.4, 0.65, 12.5, 0.35, 9, color=ORANGE_RGB)

    # Table header
    headers = ["Rang", "Client", "Fonds", "Type", "Région", "AUM Financé", "Commercial"]
    col_widths = [0.6, 2.8, 2.0, 1.4, 1.2, 1.5, 1.8]
    col_x = [0.3]
    for w in col_widths[:-1]:
        col_x.append(col_x[-1] + w)

    header_y = 1.05
    for hi, (hdr, cx, cw) in enumerate(zip(headers, col_x, col_widths)):
        _rect(s2, cx, header_y, cw, 0.4, MARINE_RGB)
        _text(s2, hdr, cx+0.05, header_y+0.06, cw-0.1, 0.28,
              8, bold=True, color=BLANC_RGB)

    # Table rows
    df_funded = pipeline_df[pipeline_df["statut"] == "Funded"].sort_values(
        "funded_aum", ascending=False).head(8) if not pipeline_df.empty else pd.DataFrame()

    for ri, (_, row) in enumerate(df_funded.iterrows()):
        ry = header_y + 0.4 + ri * 0.52
        bg = LIGHT_RGB if ri % 2 == 0 else BLANC_RGB
        client_name = ("Client {:02d}".format(ri+1) if mode_comex
                       else str(row.get("nom_client","—"))[:22])
        row_vals = [
            str(ri + 1),
            client_name,
            str(row.get("fonds","—"))[:16],
            str(row.get("type_client","—")),
            str(row.get("region","—")),
            fmt_m(float(row.get("funded_aum", 0))),
            str(row.get("sales_owner","—"))[:14],
        ]
        for vi, (val, cx, cw) in enumerate(zip(row_vals, col_x, col_widths)):
            _rect(s2, cx, ry, cw, 0.48, bg)
            clr = CIEL_RGB if vi == 5 else MARINE_RGB
            fw  = True  if vi == 5 else False
            _text(s2, val, cx+0.05, ry+0.1, cw-0.1, 0.3, 9, bold=fw, color=clr)

    if df_funded.empty:
        _text(s2, "Aucun deal Funded enregistré.", 0.4, 1.6, 12.5, 0.4, 11, color=GREY_RGB)

    _rect(s2, 0, 7.1, 13.33, 0.4, LIGHT_RGB)
    _text(s2, "Document à usage interne — Confidentiel — Amundi Asset Management",
          0.5, 7.15, 12.0, 0.3, 8, color=GREY_RGB, align=PP_ALIGN.CENTER)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# TIMEFRAME HELPER
# ---------------------------------------------------------------------------
from datetime import date as _date_cls

TIMEFRAMES = ["Max", "YTD", "1M Rolling", "3M Rolling", "1Y Rolling"]

def _timeframe_cutoff(timeframe: str):
    today = _date_cls.today()
    if timeframe == "YTD":         return _date_cls(today.year, 1, 1)
    if timeframe == "1M Rolling":  return today - timedelta(days=30)
    if timeframe == "3M Rolling":  return today - timedelta(days=91)
    if timeframe == "1Y Rolling":  return today - timedelta(days=365)
    return None

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
st.set_page_config(page_title="CRM - Asset Management",
                   layout="wide", initial_sidebar_state="expanded")

# ---------------------------------------------------------------------------
# CSS (conservé intégralement)
# ---------------------------------------------------------------------------
st.markdown("""
<style>
.stApp, .main .block-container {
    background-color: #ffffff;
    color: #001c4b;
    font-family: 'Segoe UI', Arial, sans-serif;
}
[data-testid="stSidebar"] { background-color: #001c4b; }
[data-testid="stSidebar"] * { color: #ffffff !important; }
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 { color: #019ee1 !important; }
.crm-header {
    background: #001c4b; padding: 16px 24px;
    margin-bottom: 16px; border-bottom: 3px solid #019ee1;
}
.crm-header h1 { color:#ffffff !important; margin:0; font-size:1.4rem; font-weight:700; }
.crm-header p  { color:#7ab8d8; margin:3px 0 0 0; font-size:0.80rem; }

a.clink {
    display: block; text-decoration: none !important;
    color: inherit; cursor: pointer; outline: none;
}
a.clink:hover, a.clink:focus, a.clink:visited { text-decoration: none !important; color: inherit; }

.kpi-grid {
    display: grid;
    grid-template-columns: repeat(5, 1fr);
    gap: 8px; margin-bottom: 4px;
}
.kpi-card {
    background: #001c4b; padding: 14px 10px 13px 10px;
    text-align: center; border: 1px solid #1a5e8a;
    border-bottom: 2px solid #019ee1;
    transition: background 0.15s, border-color 0.15s, box-shadow 0.15s;
    position: relative;
}
.kpi-card-clickable::after {
    content: "▸"; position: absolute; bottom: 5px; right: 7px;
    font-size: 0.54rem; color: #4a8fbd; opacity: 0.7;
    transition: color 0.15s, opacity 0.15s;
}
a.clink:hover .kpi-card-clickable {
    background: #0b3060; border-color: #f07d00 !important;
    box-shadow: 0 0 0 2px #f07d0050;
}
a.clink:hover .kpi-card-clickable::after { color: #f07d00; opacity: 1; }
.kpi-card-static { border-bottom: 2px solid #1a5e8a; }

.kpi-label { font-size:0.64rem; color:#7ab8d8; text-transform:uppercase; letter-spacing:0.8px; margin-bottom:5px; font-weight:600; }
.kpi-value { font-size:1.4rem; font-weight:800; color:#ffffff; }
.kpi-sub   { font-size:0.61rem; color:#c8dde8; margin-top:3px; }

.statut-grid { display: grid; gap: 6px; margin-bottom: 8px; }
.statut-pill {
    padding: 8px 6px; text-align: center;
    transition: filter 0.12s, box-shadow 0.12s; position: relative;
}
a.clink:hover .statut-pill { filter: brightness(1.10); box-shadow: 0 0 0 2px #f07d0055; }

.alert-overdue {
    background: #fef6f0; border-left: 3px solid #f07d00;
    padding: 8px 12px; margin: 3px 0;
    font-size: 0.77rem; color: #001c4b;
    transition: background 0.12s, box-shadow 0.12s; position: relative;
}
a.clink:hover .alert-overdue { background: #fdebd5; box-shadow: inset 3px 0 0 #f07d00; }

.activity-row {
    border-left: 2px solid #019ee1; padding: 7px 11px; margin: 3px 0;
    background: #f9fbfd; font-size: 0.78rem;
    transition: background 0.12s, border-left-color 0.12s; display: block;
}
a.clink:hover .activity-row { background: #e4f1f9; border-left-color: #f07d00; }

.badge-retard {
    display:inline-block; background:#f07d00; color:#ffffff;
    padding:1px 7px; font-size:0.66rem; font-weight:700; letter-spacing:0.4px;
}
.perimetre-badge {
    display:inline-block; background:#019ee122; border:1px solid #019ee155;
    padding:2px 8px; font-size:0.70rem; color:#ffffff; font-weight:600; margin:2px;
}
.detail-panel {
    background:#f2f6fa; border-left:3px solid #019ee1;
    padding:14px 16px 11px 16px; margin-top:10px;
}
.sales-card { background:#f4f8fc; border:1px solid #001c4b18; padding:13px; border-top:3px solid #019ee1; }
.sales-card-name { font-size:0.87rem; font-weight:700; color:#001c4b; margin-bottom:8px; padding-bottom:5px; border-bottom:1px solid #e8e8e8; }
.sales-metric { font-size:0.69rem; color:#666; margin-bottom:2px; }
.sales-metric-val { font-size:0.93rem; font-weight:700; color:#001c4b; }
.sales-metric-acc { font-size:0.93rem; font-weight:700; color:#019ee1; }
.section-title { font-size:0.90rem; font-weight:700; color:#001c4b; border-bottom:2px solid #001c4b22; padding-bottom:4px; margin:14px 0 9px 0; }
.pipeline-hint { background:#019ee108; border-left:2px solid #001c4b; padding:6px 11px; font-size:0.77rem; color:#001c4b; margin-bottom:8px; }
.sidebar-kpi { background:#ffffff14; padding:8px; margin-bottom:5px; }

.stTabs [data-baseweb="tab-list"] { background:#f0f4f8; border-bottom:2px solid #001c4b20; gap:0; }
.stTabs [data-baseweb="tab"] { color:#001c4b; font-weight:600; font-size:0.81rem; padding:7px 16px; background:#f0f4f8; border-right:1px solid #d0d8e0; }
.stTabs [aria-selected="true"] { background:#001c4b !important; color:#ffffff !important; }

.stButton > button {
    background:#019ee1; color:#ffffff; border:none; font-weight:600;
    padding:6px 15px; font-size:0.80rem; transition:background 0.12s;
}
.stButton > button:hover { background:#f07d00 !important; color:#ffffff !important; }
[data-testid="stDownloadButton"] > button { background:#019ee1 !important; color:#ffffff !important; border:none !important; font-weight:600 !important; }
[data-testid="stDownloadButton"] > button:hover { background:#f07d00 !important; color:#ffffff !important; }
[data-testid="stFormSubmitButton"] > button { background:#001c4b !important; color:#ffffff !important; font-weight:700 !important; }
[data-testid="stFormSubmitButton"] > button:hover { background:#019ee1 !important; }
.stSelectbox label, .stTextInput label, .stNumberInput label,
.stDateInput label, .stTextArea label, .stRadio label { color:#001c4b !important; font-weight:600; font-size:0.78rem; }
h1,h2,h3,h4 { color:#001c4b !important; }
hr { border-color:#001c4b10; }
code { background:#001c4b08; color:#001c4b; }

.statut-badge {
    display: inline-block; padding: 2px 8px; font-size: 0.68rem;
    font-weight: 700; border-radius: 0; letter-spacing: 0.3px;
}
.statut-Funded        { background:#1a5e8a22; color:#1a5e8a; border:1px solid #1a5e8a55; }
.statut-Soft-Commit   { background:#2c7fb822; color:#2c7fb8; border:1px solid #2c7fb855; }
.statut-Due-Diligence { background:#004f8c22; color:#004f8c; border:1px solid #004f8c55; }
.statut-Initial-Pitch { background:#4a8fbd22; color:#4a8fbd; border:1px solid #4a8fbd55; }
.statut-Prospect      { background:#9ecae122; color:#2c7fb8; border:1px solid #9ecae155; }
.statut-Lost          { background:#88888822; color:#555;    border:1px solid #aaaaaa55; }
.statut-Paused        { background:#c0c0c022; color:#666;    border:1px solid #c0c0c055; }
.statut-Redeemed      { background:#b8b8d022; color:#555;    border:1px solid #b8b8d055; }


/* ── Quick Action Center — force Marine primary buttons ── */
[data-testid="stButton"][key^="qac_"] > button,
button[kind="primary"] {
    background: #001c4b !important;
    color: #ffffff !important;
    border: none !important;
    font-weight: 700 !important;
}
button[kind="primary"]:hover {
    background: #019ee1 !important;
    color: #ffffff !important;
}

/* ── Structure card in CRM split-screen ── */
.struct-block {
    background:#f4f8fc; border-left:3px solid #001c4b30;
    padding:10px 14px; margin:6px 0; font-size:0.78rem;
}
.struct-label { font-size:0.65rem; color:#888; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:3px; }
.struct-val   { font-weight:700; color:#001c4b; }
</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# INIT DB
# ---------------------------------------------------------------------------
@st.cache_resource
def _init():
    db.init_db()
    return True
_init()


# ---------------------------------------------------------------------------
# VISUAL HELPERS
# ---------------------------------------------------------------------------
def statut_badge(statut):
    css = "statut-" + statut.replace(" ", "-")
    return '<span class="statut-badge {}">{}</span>'.format(css, statut)

def statut_dot(statut):
    dot_colors = {
        "Funded":"#1a5e8a","Soft Commit":"#2c7fb8","Due Diligence":"#004f8c",
        "Initial Pitch":"#4a8fbd","Prospect":"#9ecae1","Lost":"#aaaaaa",
        "Paused":"#c0c0c0","Redeemed":"#b8b8d0",
    }
    c = dot_colors.get(statut, "#888")
    return ('<span style="display:inline-block;width:8px;height:8px;border-radius:50%;'
            'background:{};margin-right:5px;vertical-align:middle;"></span>').format(c)

def _kyc_dot(kyc):
    colors = {"Validé": "#22a062", "En cours": ORANGE, "Bloqué": "#c0392b"}
    c = colors.get(kyc, "#aaa")
    return ('<span style="display:inline-block;width:10px;height:10px;border-radius:50%;'
            'background:{c};vertical-align:middle;margin-right:6px;" '
            'title="KYC : {kyc}"></span>').format(c=c, kyc=kyc)

def _tier_badge(tier):
    tier_colors = {"Tier 1": MARINE, "Tier 2": B_MID, "Tier 3": B_PAL}
    c = tier_colors.get(tier, GRIS)
    return ('<span style="background:{c};color:#fff;font-size:0.62rem;font-weight:700;'
            'padding:2px 8px;letter-spacing:0.3px;vertical-align:middle;">{t}</span>'
            ).format(c=c, t=tier)

def _content_funded(fonds_filter=None):
    df = db.get_funded_deals_detail(fonds_filter)
    if df.empty: st.info("Aucun deal Funded."); return
    df2 = df.copy()
    for col in ["AUM_Finance","AUM_Cible"]:
        if col in df2.columns: df2[col] = df2[col].apply(fmt_m)
    st.dataframe(df2, use_container_width=True, hide_index=True, height=min(380, 46 + len(df2) * 36))

def _content_pipeline(fonds_filter=None):
    df = db.get_active_deals_detail(fonds_filter)
    if df.empty: st.info("Aucun deal actif."); return
    df2 = df.copy()
    # Show Smart AUM as single "AUM Attendu" column
    if "AUM_Pipeline" in df2.columns:
        df2["AUM Attendu"] = df2["AUM_Pipeline"].apply(fmt_m)
        drop_cols = [c for c in ["AUM_Revise","AUM_Cible","AUM_Pipeline"] if c in df2.columns]
        df2 = df2.drop(columns=drop_cols)
    if "Prochaine_Action" in df2.columns:
        df2["Prochaine_Action"] = df2["Prochaine_Action"].apply(
            lambda d: d.isoformat() if isinstance(d, date) else str(d or "—"))
    st.dataframe(df2, use_container_width=True, hide_index=True, height=min(380, 46 + len(df2) * 36))

def _content_lost(fonds_filter=None):
    df = db.get_lost_deals_detail(fonds_filter)
    if df.empty: st.info("Aucun deal Lost/Paused."); return
    df2 = df.copy()
    if "AUM_Cible" in df2.columns: df2["AUM_Cible"] = df2["AUM_Cible"].apply(fmt_m)
    st.dataframe(df2, use_container_width=True, hide_index=True, height=min(380, 46 + len(df2) * 36))

def _content_overdue():
    df = db.get_overdue_actions()
    if df.empty: st.info("Aucune action en retard."); return
    today = date.today()
    for _, row in df.iterrows():
        pid  = int(row.get("id", 0)) if "id" in df.columns else None
        deal = db.get_overdue_deal_full(pid) if pid else None
        nad  = row.get("next_action_date")
        days = (today - nad).days if isinstance(nad, date) else 0
        nad_str = nad.isoformat() if isinstance(nad, date) else "—"
        st.markdown(
            "<div style='border-left:3px solid #f07d00;padding:8px 14px;"
            "margin:8px 0;background:#fef6f0;'>"
            "<div style='font-size:0.84rem;font-weight:700;color:#001c4b;'>"
            "{} &nbsp;<span style='background:#f07d00;color:#fff;padding:1px 7px;"
            "font-size:0.66rem;font-weight:700;'>RETARD +{}j</span></div>".format(
                row.get("nom_client",""), days), unsafe_allow_html=True)
        if deal:
            c1, c2, c3 = st.columns(3)
            with c1:
                st.markdown("**Fonds**"); st.write(deal.get("fonds","—"))
                st.markdown("**Statut**"); st.write(deal.get("statut","—"))
                st.markdown("**Commercial**"); st.write(deal.get("sales_owner","—"))
            with c2:
                st.markdown("**AUM Cible**"); st.write(fmt_m(deal.get("target_aum_initial",0)))
                st.markdown("**AUM Revise**"); st.write(fmt_m(deal.get("revised_aum",0)))
                st.markdown("**AUM Finance**"); st.write(fmt_m(deal.get("funded_aum",0)))
            with c3:
                st.markdown("**Region**"); st.write(deal.get("region","—"))
                st.markdown("**Prochaine Action**"); st.write(nad_str)
                st.markdown("**Derniere Activite**")
                st.write(deal.get("derniere_activite","") or "—")
        st.markdown("</div>", unsafe_allow_html=True)

def _content_statut(statut_nom, fonds_filter=None):
    df = db.get_pipeline_by_statut(statut_nom, fonds_filter)
    if df.empty: st.info("Aucun deal en statut {}.".format(statut_nom)); return
    today = date.today()
    for _, row in df.iterrows():
        nad = row.get("next_action_date")
        if isinstance(nad, date):
            delta = (nad - today).days
            nad_str = nad.isoformat()
            timing_html = (
                "<span class='badge-retard'>RETARD +{}j</span>".format(abs(delta)) if delta < 0
                else "<span style='color:#019ee1;font-weight:700;'>Aujourd'hui</span>" if delta == 0
                else "<span style='color:#444;'>Dans {}j</span>".format(delta))
        else:
            nad_str = "—"; timing_html = "—"
        act = str(row.get("derniere_activite",""))
        st.markdown(
            "<div style='border-left:3px solid {ciel};padding:8px 14px;"
            "margin:6px 0;background:#f4f8fc;'>"
            "<div style='font-size:0.85rem;font-weight:700;color:{marine};'>{client}</div>"
            "<div style='font-size:0.76rem;color:#555;margin:2px 0;'>"
            "{fonds} &middot; {type} &middot; {region}</div>"
            "<div style='display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;margin-top:6px;'>"
            "<div><div style='font-size:0.67rem;color:#888;'>AUM Revise</div>"
            "<div style='font-weight:700;color:{marine};font-size:0.82rem;'>{aum}</div></div>"
            "<div><div style='font-size:0.67rem;color:#888;'>Prochaine Action</div>"
            "<div style='font-size:0.78rem;'>{nad}&nbsp;{timing}</div></div>"
            "<div><div style='font-size:0.67rem;color:#888;'>Commercial</div>"
            "<div style='font-size:0.78rem;color:#444;'>{owner}</div></div>"
            "</div>"
            "{act_block}</div>".format(
                ciel=CIEL, marine=MARINE, client=row.get("nom_client",""),
                fonds=row.get("fonds",""), type=row.get("type_client",""),
                region=row.get("region",""),
                aum=fmt_m(float(row.get("aum_pipeline", 0) or 0) if row.get("aum_pipeline") is not None
                        else (float(row.get("revised_aum",0) or 0) if float(row.get("revised_aum",0) or 0) > 0
                              else float(row.get("target_aum_initial",0) or 0))),
                nad=nad_str, timing=timing_html, owner=row.get("sales_owner",""),
                act_block=(
                    "<div style='margin-top:5px;font-size:0.73rem;color:#666;"
                    "border-top:1px solid #e8e8e8;padding-top:4px;'>"
                    "<b>Derniere activite :</b> {}</div>".format(act)
                ) if act else ""), unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# STATE GUARD — prevents cross-tab state pollution
# Each tab uses its own namespace prefix in session_state
# ---------------------------------------------------------------------------
_filtre_effectif = None  # set in sidebar


# ============================================================================
# DIALOG FACTORY — all modals in one place
# ============================================================================

@st.dialog("Ajouter un compte client", width="large")
def dialog_add_client():
    with st.form("dlg_form_client", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            nom_client  = st.text_input("Nom du Client")
            type_client = st.selectbox("Type Client", TYPES_CLIENT)
            region      = st.selectbox("Région", REGIONS)
        with c2:
            is_parent = st.checkbox("Ce compte est une Maison Mère", value=False)
            all_clients_opts = db.get_all_clients()
            parent_options   = ["(Aucune — client autonome)"] + all_clients_opts["nom_client"].tolist()
            parent_sel       = st.selectbox("Maison Mère (optionnel)", parent_options,
                                             disabled=is_parent)
            tier_sel    = st.selectbox("Tier", db.TIERS_REFERENTIEL, index=1)
            kyc_sel     = st.selectbox("Statut KYC", db.KYC_STATUTS, index=1)
            country_sel = st.selectbox("Country", COUNTRIES_LIST)
        interests_sel = st.multiselect("Product Interests", db.PRODUCT_INTERESTS)
        if st.form_submit_button("Enregistrer le compte", use_container_width=True):
            if not nom_client.strip():
                st.error("Nom obligatoire.")
            else:
                try:
                    _parent_id = None
                    if not is_parent and parent_sel != "(Aucune — client autonome)":
                        _pr = all_clients_opts[all_clients_opts["nom_client"] == parent_sel]
                        if not _pr.empty:
                            _parent_id = int(_pr.iloc[0]["id"])
                    db.add_client(nom_client.strip(), type_client, region,
                                  country=country_sel, parent_id=_parent_id,
                                  tier=tier_sel, kyc_status=kyc_sel,
                                  product_interests=",".join(interests_sel))
                    st.success("Client {} ajouté.".format(nom_client))
                    st.rerun()
                except Exception as e:
                    st.warning("Ce client existe déjà." if "UNIQUE" in str(e)
                               else "Erreur : {}".format(e))


@st.dialog("Modifier le compte", width="large")
def dialog_edit_client(client_data: dict):
    sel_id   = int(client_data["id"])
    sel_nom  = str(client_data.get("nom_client",""))
    sel_type = str(client_data.get("type_client",""))
    sel_reg  = str(client_data.get("region",""))
    sel_tier = str(client_data.get("tier","Tier 2"))
    sel_kyc  = str(client_data.get("kyc_status","En cours"))
    sel_par  = str(client_data.get("parent_nom",""))
    cur_country = str(client_data.get("country",""))
    cur_prod    = [p.strip() for p in str(client_data.get("product_interests","")).split(",") if p.strip()]

    all_clients_edit = db.get_all_clients()
    parent_opts_edit = ["(Aucune — client autonome)"] + [
        n for n in all_clients_edit["nom_client"].tolist() if n != sel_nom]
    cur_parent_nom = sel_par or "(Aucune — client autonome)"

    with st.form("dlg_edit_company_{}".format(sel_id), clear_on_submit=False):
        ecp1, ecp2 = st.columns(2)
        with ecp1:
            ec_nom    = st.text_input("Nom du compte", value=sel_nom)
            ec_type   = st.selectbox("Type Client", TYPES_CLIENT,
                index=TYPES_CLIENT.index(sel_type) if sel_type in TYPES_CLIENT else 0)
            ec_region = st.selectbox("Région", REGIONS,
                index=REGIONS.index(sel_reg) if sel_reg in REGIONS else 0)
            ec_country = st.selectbox("Country", COUNTRIES_LIST,
                index=COUNTRIES_LIST.index(cur_country) if cur_country in COUNTRIES_LIST else 0)
        with ecp2:
            ec_tier   = st.selectbox("Tier", db.TIERS_REFERENTIEL,
                index=db.TIERS_REFERENTIEL.index(sel_tier) if sel_tier in db.TIERS_REFERENTIEL else 1)
            ec_kyc    = st.selectbox("Statut KYC", db.KYC_STATUTS,
                index=db.KYC_STATUTS.index(sel_kyc) if sel_kyc in db.KYC_STATUTS else 1)
            ec_parent = st.selectbox("Maison Mère", parent_opts_edit,
                index=parent_opts_edit.index(cur_parent_nom)
                      if cur_parent_nom in parent_opts_edit else 0)
            ec_is_parent = st.checkbox("Ce compte est une Maison Mère", value=(sel_par == ""))
        ec_interests = st.multiselect("Product Interests", db.PRODUCT_INTERESTS, default=cur_prod)
        esave, ecancel = st.columns(2)
        with esave:
            if st.form_submit_button("Enregistrer les modifications", use_container_width=True):
                _new_parent_id = None
                if not ec_is_parent and ec_parent != "(Aucune — client autonome)":
                    _pr = all_clients_edit[all_clients_edit["nom_client"] == ec_parent]
                    if not _pr.empty:
                        _new_parent_id = int(_pr.iloc[0]["id"])
                ok, err = db.update_client(
                    sel_id, ec_nom.strip(), ec_type, ec_region,
                    country=ec_country, parent_id=_new_parent_id,
                    tier=ec_tier, kyc_status=ec_kyc,
                    product_interests=",".join(ec_interests))
                if ok:
                    st.success("Compte mis à jour.")
                    st.rerun()
                else:
                    st.error("Erreur : {}".format(err))
        with ecancel:
            if st.form_submit_button("Annuler", use_container_width=True):
                st.rerun()


@st.dialog("Ajouter un contact", width="large")
def dialog_add_contact(client_id: int, client_nom: str = ""):
    st.caption("Compte : **{}**".format(client_nom) if client_nom else "")
    with st.form("dlg_add_contact_{}".format(client_id), clear_on_submit=True):
        cc1, cc2 = st.columns(2)
        with cc1:
            qc_prenom = st.text_input("Prénom")
            qc_nom    = st.text_input("Nom")
            qc_role   = st.selectbox("Rôle", [""] + db.ROLES_CONTACT)
        with cc2:
            qc_email   = st.text_input("Email")
            qc_tel     = st.text_input("Téléphone")
            qc_li      = st.text_input("LinkedIn (URL)")
        qc_primary = st.checkbox("Contact principal")
        if st.form_submit_button("Enregistrer le Contact", use_container_width=True):
            if not qc_nom.strip():
                st.error("Nom obligatoire.")
            else:
                db.add_contact(client_id, qc_prenom, qc_nom, qc_role,
                               qc_email, qc_tel, qc_li, qc_primary)
                st.success("Contact ajouté.")
                st.rerun()


@st.dialog("Modifier le contact", width="large")
def dialog_edit_contact(contact_data: dict):
    ct_id = int(contact_data.get("id", 0))
    with st.form("dlg_edit_ct_{}".format(ct_id), clear_on_submit=True):
        ec1, ec2 = st.columns(2)
        with ec1:
            ec_prenom = st.text_input("Prénom", value=str(contact_data.get("prenom","")))
            ec_nom    = st.text_input("Nom",    value=str(contact_data.get("nom","")))
            ec_role   = st.selectbox("Rôle", [""] + db.ROLES_CONTACT,
                index=([""] + db.ROLES_CONTACT).index(str(contact_data.get("role","")))
                      if str(contact_data.get("role","")) in ([""] + db.ROLES_CONTACT) else 0)
        with ec2:
            ec_email = st.text_input("Email",     value=str(contact_data.get("email","")))
            ec_tel   = st.text_input("Téléphone", value=str(contact_data.get("telephone","")))
            ec_li    = st.text_input("LinkedIn",  value=str(contact_data.get("linkedin","")))
        ec_primary = st.checkbox("Contact principal", value=bool(contact_data.get("is_primary", 0)))
        ecs1, ecs2 = st.columns(2)
        with ecs1:
            if st.form_submit_button("Enregistrer les modifications", use_container_width=True):
                ok, err = db.update_contact(ct_id, ec_prenom, ec_nom,
                                            ec_role, ec_email, ec_tel, ec_li, ec_primary)
                if ok:
                    st.success("Contact mis à jour.")
                    st.rerun()
                else:
                    st.error(err)
        with ecs2:
            if st.form_submit_button("Annuler", use_container_width=True):
                st.rerun()


@st.dialog("Ajouter un deal", width="large")
def dialog_add_deal(preselect_client_id: int = None):
    _clients = db.get_client_options()
    _st_team = db.get_sales_team()
    _marchés = sorted(_st_team["marche"].unique().tolist()) if not _st_team.empty else ["Global"]
    if not _clients:
        st.info("Ajoutez d'abord un client.")
        return
    with st.form("dlg_add_deal", clear_on_submit=True):
        ca, cb = st.columns(2)
        with ca:
            # Pre-select client if called from CRM tab
            client_names = list(_clients.keys())
            default_idx = 0
            if preselect_client_id:
                _matching = [i for i, cid in enumerate(_clients.values())
                             if cid == preselect_client_id]
                if _matching:
                    default_idx = _matching[0]
            client_sel  = st.selectbox("Client", client_names, index=default_idx)
            fonds_sel   = st.selectbox("Fonds", FONDS)
            statut_sel  = st.selectbox("Statut", STATUTS)
            marche_sel  = st.selectbox("Marché", ["Tous"] + _marchés)
            if marche_sel == "Tous":
                _owners = _st_team["nom"].tolist() if not _st_team.empty else ["Non assigne"]
            else:
                _owners = _st_team[_st_team["marche"]==marche_sel]["nom"].tolist()
            if not _owners: _owners = ["Non assigne"]
            owner_sel = st.selectbox("Commercial", _owners)
        with cb:
            target_aum_m  = st.number_input("AUM Cible (en M€)", min_value=0.0, step=1.0, format="%.2f")
            revised_aum_m = st.number_input("AUM Révisé (en M€)", min_value=0.0, step=1.0, format="%.2f")
            funded_aum_m  = st.number_input("AUM Financé (en M€)", min_value=0.0, step=1.0, format="%.2f",
                                             help="Laisser à 0 si Funded → utilise l'AUM Révisé automatiquement")
            closing_prob = st.slider("Probabilité de closing (%)", 0, 100, 50, 5)
        raison_perte, concurrent = "", ""
        if statut_sel in ("Lost","Paused"):
            cc, cd = st.columns(2)
            with cc: raison_perte = st.selectbox("Raison", RAISONS_PERTE)
            with cd: concurrent   = st.text_input("Concurrent")
        next_action = st.date_input("Prochaine Action", value=date.today() + timedelta(days=14))
        if st.form_submit_button("Enregistrer le Deal", use_container_width=True):
            if statut_sel in ("Lost","Paused") and not raison_perte:
                st.error("Raison obligatoire.")
            else:
                db.add_pipeline_entry(
                    _clients[client_sel], fonds_sel, statut_sel,
                    target_aum_m * 1_000_000,
                    revised_aum_m * 1_000_000,
                    funded_aum_m * 1_000_000,
                    raison_perte, concurrent, next_action.isoformat(),
                    owner_sel, closing_prob)
                st.success("Deal {} / {} enregistré.".format(fonds_sel, client_sel))
                st.rerun()


@st.dialog("Enregistrer une activité", width="large")
def dialog_add_activity(preselect_client_id: int = None):
    _clients = db.get_client_options()
    if not _clients:
        st.info("Ajoutez d'abord un client.")
        return
    with st.form("dlg_add_activity", clear_on_submit=True):
        client_names = list(_clients.keys())
        default_idx  = 0
        if preselect_client_id:
            _matching = [i for i, cid in enumerate(_clients.values())
                         if cid == preselect_client_id]
            if _matching:
                default_idx = _matching[0]
        dlg_client = st.selectbox("Client", client_names, index=default_idx)
        dlg_type   = st.selectbox("Type d'interaction", TYPES_INTERACTION)
        dlg_date   = st.date_input("Date", value=date.today())
        dlg_notes  = st.text_area("Notes", height=100)
        if st.form_submit_button("Enregistrer", use_container_width=True):
            db.add_activity(_clients[dlg_client], dlg_date.isoformat(), dlg_notes, dlg_type)
            st.success("Activité enregistrée.")
            st.rerun()


@st.dialog("Modifier le Deal", width="large")
def dialog_edit_pipeline(pipeline_id: int, row_data: dict):
    """Modale d'édition pipeline — sans scroll, état isolé via clé 'crm_pipeline_dialog_id'."""
    client_name    = str(row_data.get("nom_client",""))
    current_statut = str(row_data.get("statut","Prospect"))

    st.markdown(
        '<div style="font-size:0.84rem;font-weight:700;color:{marine};">'
        'Client : <span style="color:{ciel};">{name}</span>'
        '&nbsp;{badge}&nbsp;<span style="font-size:0.68rem;color:#888;">ID #{pid}</span>'
        '</div>'.format(
            marine=MARINE, ciel=B_MID, name=client_name,
            badge=statut_badge(current_statut), pid=pipeline_id),
        unsafe_allow_html=True)

    _st_team    = db.get_sales_team()
    _marchés_ed = sorted(_st_team["marche"].unique().tolist()) if not _st_team.empty else ["Global"]

    with st.form(key="dlg_edit_pipeline_{}".format(pipeline_id)):
        r1c1, r1c2 = st.columns(2)
        with r1c1:
            fi = FONDS.index(row_data["fonds"]) if row_data["fonds"] in FONDS else 0
            new_fonds  = st.selectbox("Fonds", FONDS, index=fi)
            cur_owner  = str(row_data.get("sales_owner") or "Non assigne")
            cur_marche = "Tous"
            if not _st_team.empty:
                rows_m = _st_team[_st_team["nom"] == cur_owner]
                if not rows_m.empty:
                    cur_marche = rows_m.iloc[0]["marche"]
            marche_edit = st.selectbox(
                "Marché commercial", ["Tous"] + _marchés_ed,
                index=(["Tous"] + _marchés_ed).index(cur_marche)
                      if cur_marche in (["Tous"] + _marchés_ed) else 0,
                key="dlg_marche_{}".format(pipeline_id))
            if marche_edit == "Tous":
                _owners_edit = _st_team["nom"].tolist() if not _st_team.empty else ["Non assigne"]
            else:
                _owners_edit = _st_team[_st_team["marche"]==marche_edit]["nom"].tolist()
            if not _owners_edit: _owners_edit = ["Non assigne"]
            oi = _owners_edit.index(cur_owner) if cur_owner in _owners_edit else 0
            new_sales = st.selectbox("Commercial", _owners_edit, index=oi,
                                     key="dlg_owner_{}".format(pipeline_id))
        with r1c2:
            si = STATUTS.index(current_statut) if current_statut in STATUTS else 0
            new_statut = st.selectbox("Statut", STATUTS, index=si)
            cur_prob   = float(row_data.get("closing_probability") or 50)
            new_prob   = st.slider("Probabilité de closing (%)", 0, 100, int(cur_prob), 5,
                                   key="dlg_prob_{}".format(pipeline_id))
        r2c1, r2c2, r2c3 = st.columns(3)
        with r2c1:
            new_target_m = st.number_input(
                "AUM Cible (en M€)",
                value=round(float(row_data.get("target_aum_initial",0.0)) / 1_000_000, 2),
                min_value=0.0, step=1.0, format="%.2f")
        with r2c2:
            new_revised_m = st.number_input(
                "AUM Révisé (en M€)",
                value=round(float(row_data.get("revised_aum",0.0)) / 1_000_000, 2),
                min_value=0.0, step=1.0, format="%.2f")
        with r2c3:
            new_funded_m = st.number_input(
                "AUM Financé (en M€)",
                value=round(float(row_data.get("funded_aum",0.0)) / 1_000_000, 2),
                min_value=0.0, step=1.0, format="%.2f",
                help="Laisser à 0 si Funded → utilise l'AUM Révisé automatiquement")
        r3c1, r3c2, r3c3 = st.columns(3)
        with r3c1:
            ropts  = [""] + RAISONS_PERTE
            cur_r  = str(row_data.get("raison_perte") or "")
            ri     = ropts.index(cur_r) if cur_r in ropts else 0
            lbl_r  = "Raison (obligatoire)" if new_statut in ("Lost","Paused") else "Raison"
            new_raison = st.selectbox(lbl_r, ropts, index=ri)
        with r3c2:
            new_conc = st.text_input("Concurrent", value=str(row_data.get("concurrent_choisi") or ""))
        with r3c3:
            nad = row_data.get("next_action_date")
            if not isinstance(nad, date): nad = date.today() + timedelta(days=14)
            new_nad = st.date_input("Prochaine Action", value=nad)

        col_save, col_cancel = st.columns([3, 1])
        with col_save:
            sub    = st.form_submit_button("Sauvegarder les modifications", use_container_width=True)
        with col_cancel:
            cancel = st.form_submit_button("Annuler", use_container_width=True)

    # Audit log inside dialog
    df_audit = db.get_audit_log(pipeline_id)
    if not df_audit.empty:
        with st.expander("Historique modifications", expanded=False):
            st.dataframe(df_audit, use_container_width=True, hide_index=True,
                         height=min(200, 46 + len(df_audit) * 36))

    # Delete deal — inside dialog
    st.markdown("---")
    if st.button("Supprimer ce deal (irréversible)", type="tertiary",
                 key="dlg_del_btn_{}".format(pipeline_id)):
        st.session_state["crm_pipeline_del_confirm"] = pipeline_id

    if st.session_state.get("crm_pipeline_del_confirm") == pipeline_id:
        st.error("Suppression définitive du deal #{} et de son historique.".format(pipeline_id))
        dd1, dd2, _ = st.columns([1, 1, 4])
        with dd1:
            if st.button("Confirmer", key="dlg_del_yes_{}".format(pipeline_id)):
                ok, err = db.delete_pipeline_row(pipeline_id)
                if ok:
                    st.session_state.pop("crm_pipeline_del_confirm", None)
                    st.session_state.pop("crm_pipeline_dialog_id", None)
                    st.rerun()
                else:
                    st.error(err)
        with dd2:
            if st.button("Annuler", key="dlg_del_no_{}".format(pipeline_id)):
                st.session_state.pop("crm_pipeline_del_confirm", None)
                st.rerun()

    # Form submission
    if cancel:
        st.session_state.pop("crm_pipeline_dialog_id", None)
        st.rerun()

    if sub:
        new_target  = new_target_m  * 1_000_000
        new_revised = new_revised_m * 1_000_000
        new_funded  = new_funded_m  * 1_000_000
        if new_statut == "Funded":
            kyc_check = str(row_data.get("kyc_status", "En cours"))
            if kyc_check in ("Bloqué", "En cours"):
                st.error("KYC incomplet ({}). Validez le KYC avant de passer au statut Funded.".format(kyc_check))
                return
        ok, msg = db.update_pipeline_row({
            "id": pipeline_id, "fonds": new_fonds, "statut": new_statut,
            "target_aum_initial": new_target, "revised_aum": new_revised,
            "funded_aum": new_funded, "raison_perte": new_raison,
            "concurrent_choisi": new_conc, "next_action_date": new_nad,
            "sales_owner": new_sales, "closing_probability": new_prob})
        if ok:
            st.session_state.pop("crm_pipeline_dialog_id", None)
            st.success("Deal mis à jour.")
            st.rerun()
        else:
            st.error(msg)


@st.dialog("Modifier l'activité", width="large")
def dialog_edit_activity(act_id: int, act_data: dict):
    with st.form("dlg_edit_act_{}".format(act_id), clear_on_submit=True):
        ea1, ea2 = st.columns(2)
        with ea1:
            ea_type = st.selectbox("Type", TYPES_INTERACTION,
                index=TYPES_INTERACTION.index(act_data.get("type","Call"))
                      if act_data.get("type") in TYPES_INTERACTION else 0)
            try:
                ea_date_val = date.fromisoformat(act_data.get("date",""))
            except Exception:
                ea_date_val = date.today()
            ea_date = st.date_input("Date", value=ea_date_val)
        with ea2:
            ea_notes = st.text_area("Notes", value=act_data.get("notes",""), height=100)
        col_save, col_cancel = st.columns(2)
        with col_save:
            if st.form_submit_button("Enregistrer les modifications", use_container_width=True):
                ok, err = db.update_activity(act_id, ea_date.isoformat(), ea_notes, ea_type)
                if ok:
                    st.rerun()
                else:
                    st.error(err)
        with col_cancel:
            if st.form_submit_button("Annuler", use_container_width=True):
                st.rerun()




@st.dialog("Meeting Brief — Analyse Compte", width="large")
def dialog_meeting_brief(client_data: dict, group_summary: dict,
                          contacts_df, activities_df, whitespace_df):
    """One-minute briefing for CRM director / relationship manager."""
    sel_nom = client_data.get("nom_client", "")
    st.markdown(
        '<div style="background:#001c4b;padding:14px 18px;margin-bottom:12px;">'
        '<div style="font-size:0.72rem;color:#7ab8d8;text-transform:uppercase;letter-spacing:1px;">'
        'Meeting Brief — Confidentiel</div>'
        '<div style="font-size:1.2rem;font-weight:800;color:#ffffff;">{}</div>'
        '</div>'.format(sel_nom), unsafe_allow_html=True)

    bc1, bc2 = st.columns(2)
    with bc1:
        st.markdown("**AUM Consolidé Groupe**")
        st.markdown(
            '<span style="font-size:1.6rem;font-weight:800;color:#001c4b;">{}</span>'.format(
                fmt_m(group_summary["aum_consolide"])), unsafe_allow_html=True)
        if group_summary["fonds_investis"]:
            st.markdown("**Fonds investis :** " + " · ".join(group_summary["fonds_investis"]))
        else:
            st.info("Aucun deal Funded à ce jour.")
    with bc2:
        st.markdown("**Contacts clés**")
        if not contacts_df.empty:
            for _, ct in contacts_df.head(3).iterrows():
                st.markdown("- **{}** {} — *{}*".format(
                    ct.get("prenom",""), ct.get("nom",""), ct.get("role","—")))
        else:
            st.caption("Aucun contact enregistré.")

    st.divider()
    st.markdown("**2 Dernières Activités**")
    if not activities_df.empty:
        for _, act in activities_df.head(2).iterrows():
            st.markdown("- `{}` **{}** — {}".format(
                str(act.get("date","")),
                str(act.get("type_interaction","")),
                str(act.get("notes",""))[:120]))
    else:
        st.caption("Aucune activité.")

    st.divider()
    st.markdown("**Opportunités Whitespace (Cross-Sell)**")
    if not whitespace_df.empty and sel_nom in whitespace_df.index:
        row = whitespace_df.loc[sel_nom]
        gaps = [f for f in row.index if pd.isna(row[f])]
        invested = [f for f in row.index if not pd.isna(row[f])]
        if gaps:
            st.markdown("Fonds **non investis** (opportunités) : " +
                        ", ".join("**{}**".format(g) for g in gaps))
        if invested:
            st.markdown("Fonds investis : " + ", ".join(invested))
    else:
        st.caption("Données whitespace indisponibles.")

    if group_summary["next_actions"]:
        st.divider()
        st.markdown("**Prochaines Actions Pipeline**")
        for act in group_summary["next_actions"][:3]:
            st.markdown("- **{}** — {} — {} — `{}`".format(
                act["fonds"], act["statut"],
                fmt_m(act["aum_pipeline"]), act["nad"]))


@st.dialog("Gérer les Commerciaux", width="large")
def dialog_manage_sales():
    st_team = db.get_sales_team()
    if not st_team.empty:
        st.dataframe(st_team[["nom","marche"]], hide_index=True, use_container_width=True)
    st.markdown("**Ajouter un commercial**")
    with st.form("dlg_add_sales", clear_on_submit=True):
        nc1, nc2 = st.columns(2)
        with nc1: new_sales_nom    = st.text_input("Nom")
        with nc2: new_sales_marche = st.selectbox("Marché", ["GCC","EMEA","APAC","Americas","Nordics","Global"])
        if st.form_submit_button("Ajouter", use_container_width=True):
            if new_sales_nom.strip():
                ok = db.add_sales_member(new_sales_nom.strip(), new_sales_marche)
                st.success("Commercial ajouté." if ok else "Nom déjà existant.")
                st.rerun()


# ---------------------------------------------------------------------------
# SIDEBAR
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown(
        '<div style="text-align:center;padding:6px 0 14px 0;">'
        '<div style="font-size:0.91rem;font-weight:800;color:#ffffff;">CRM Asset Management</div>'
        '<div style="font-size:0.65rem;color:#e8e8e8;margin-top:3px;">'
        + date.today().strftime("%d %B %Y") + '</div></div>',
        unsafe_allow_html=True)
    st.divider()

    kpis_global = db.get_kpis()
    st.markdown('<div style="color:#4a8fbd;font-size:0.67rem;font-weight:700;'
                'text-transform:uppercase;letter-spacing:1px;margin-bottom:5px;">Apercu Global</div>',
                unsafe_allow_html=True)
    for lbl, val in [("AUM Finance Total", fmt_m(kpis_global["total_funded"])),
                     ("Pipeline Actif",    fmt_m(kpis_global["pipeline_actif"]))]:
        st.markdown('<div class="sidebar-kpi">'
                    '<div style="font-size:0.62rem;color:#c8dde8;">{}</div>'
                    '<div style="font-size:1.06rem;font-weight:800;color:#ffffff;">{}</div>'
                    '</div>'.format(lbl, val), unsafe_allow_html=True)

    st.divider()
    st.markdown('<div style="color:#4a8fbd;font-size:0.67rem;font-weight:700;'
                'text-transform:uppercase;letter-spacing:1px;margin-bottom:5px;">Backup</div>',
                unsafe_allow_html=True)
    try:
        excel_bytes = db.get_excel_backup()
        st.download_button("Exporter Backup Excel", data=excel_bytes,
                           file_name="backup_crm_{}.xlsx".format(date.today().isoformat()),
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
    except Exception as e_xl:
        st.error("Backup Excel : {}".format(e_xl))

    st.divider()
    st.markdown('<div style="color:#f07d00;font-size:0.67rem;font-weight:700;'
                'text-transform:uppercase;letter-spacing:1px;margin-bottom:5px;">Export Hub</div>',
                unsafe_allow_html=True)

    fonds_perimetre = st.multiselect("Périmètre de l'export", options=FONDS, default=FONDS,
                                     key="fonds_perimetre_select")
    _filtre_effectif = (fonds_perimetre
                        if (fonds_perimetre and len(fonds_perimetre) < len(FONDS))
                        else None)
    mode_comex = st.toggle("Mode Comex — Anonymisation", value=False, key="hub_mode_comex")

    if fonds_perimetre and len(fonds_perimetre) < len(FONDS):
        st.markdown("<div>{}</div>".format(" ".join(
            '<span class="perimetre-badge">{}</span>'.format(f) for f in fonds_perimetre)),
            unsafe_allow_html=True)
    elif not fonds_perimetre:
        st.warning("Sélectionnez au moins un fonds.")

    hub_format = st.radio(
        "Format du rapport",
        ["PDF", "PPTX", "Email"],
        horizontal=True,
        key="hub_format_radio"
    )

    # PDF-only options
    if hub_format == "PDF":
        st.markdown('<div style="color:#8ab8d8;font-size:0.62rem;font-weight:600;'
                    'text-transform:uppercase;letter-spacing:0.7px;margin:6px 0 3px 0;">Sections PDF</div>',
                    unsafe_allow_html=True)
        include_top10    = st.checkbox("Top 10 Inflows",     value=True,  key="hub_top10")
        include_outflows = st.checkbox("Top 10 Outflows",    value=False, key="hub_outflows",
                                       disabled=not include_top10)
        include_perf_pdf = st.checkbox("Performance NAV",    value=True,  key="hub_perf_nav")
    else:
        include_top10 = True; include_outflows = False; include_perf_pdf = False

    perf_data_pdf   = st.session_state.get("perf_data",   None)
    nav_base100_pdf = st.session_state.get("nav_base100", None)
    if perf_data_pdf is not None and hub_format == "PDF":
        st.caption("NAV chargée : {} fonds.".format(len(perf_data_pdf)))

    # ── Single "Générer le Rapport" button ────────────────────────────────────
    _btn_disabled = not fonds_perimetre
    if _btn_disabled:
        st.button("Générer le Rapport", key="hub_gen_btn_disabled",
                  disabled=True, use_container_width=True)
    elif st.button("Générer le Rapport", key="hub_gen_btn", use_container_width=True):
        with st.spinner("Génération en cours…"):
            try:
                pipeline_hub = db.get_pipeline_with_clients(fonds_filter=_filtre_effectif)
                kpis_hub     = db.get_kpis(fonds_filter=_filtre_effectif)

                if hub_format == "PDF":
                    aum_region_pdf = db.get_aum_by_region(fonds_filter=_filtre_effectif)
                    pf_pdf = perf_data_pdf
                    nb_pdf = nav_base100_pdf
                    if pf_pdf is not None and _filtre_effectif and "Fonds" in pf_pdf.columns:
                        pf_pdf = pf_pdf[pf_pdf["Fonds"].isin(_filtre_effectif)]
                    if nb_pdf is not None and _filtre_effectif and hasattr(nb_pdf, "columns"):
                        cols_k = [c for c in nb_pdf.columns if c in _filtre_effectif]
                        nb_pdf = nb_pdf[cols_k] if cols_k else None
                    _include_perf = (include_perf_pdf and pf_pdf is not None
                                     and hasattr(pf_pdf, "empty") and not pf_pdf.empty)
                    pdf_bytes = pdf_gen.generate_pdf(
                        pipeline_df=pipeline_hub, kpis=kpis_hub,
                        aum_by_region=aum_region_pdf, mode_comex=mode_comex,
                        perf_data=pf_pdf, nav_base100_df=nb_pdf,
                        fonds_perimetre=fonds_perimetre,
                        include_top10=include_top10, include_outflows=include_outflows,
                        include_perf=_include_perf)
                    fname_pdf = "report{}_{}.pdf".format(
                        "_comex" if mode_comex else "", date.today().isoformat())
                    st.download_button("Télécharger le PDF", data=pdf_bytes,
                                       file_name=fname_pdf, mime="application/pdf",
                                       key="hub_dl_pdf", use_container_width=True)
                    st.success("PDF généré.")

                elif hub_format == "PPTX":
                    pptx_bytes = generate_global_pptx(kpis_hub, pipeline_hub,
                                                       mode_comex=mode_comex)
                    fname_pptx = "rapport_global{}_{}.pptx".format(
                        "_comex" if mode_comex else "", date.today().isoformat())
                    st.download_button(
                        "Télécharger le PPTX", data=pptx_bytes,
                        file_name=fname_pptx,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        key="hub_dl_pptx", use_container_width=True)
                    st.success("PPTX généré.")

                elif hub_format == "Email":
                    # Build email summary
                    _aum_f   = fmt_m(kpis_hub.get("total_funded", 0))
                    _pip_a   = fmt_m(kpis_hub.get("pipeline_actif", 0))
                    _wp      = fmt_m(kpis_hub.get("weighted_pipeline", 0))
                    _taux    = "{:.1f}%".format(kpis_hub.get("taux_conversion", 0))
                    _nb_f    = kpis_hub.get("nb_funded", 0)
                    _nb_l    = kpis_hub.get("nb_lost", 0)
                    _nb_a    = kpis_hub.get("nb_deals_actifs", 0)
                    _perims  = ", ".join(fonds_perimetre) if fonds_perimetre else "Tous les fonds"
                    _comex_note = " *(noms masqués — Mode Comex)*" if mode_comex else ""

                    md_summary = """
**Rapport Portfolio — {}{}**

| KPI | Valeur |
|-----|--------|
| AUM Financé Total | {} |
| Pipeline Actif (Smart AUM) | {} |
| Pipeline Pondéré | {} |
| Taux de Conversion | {} |
| Deals Funded | {} |
| Deals Actifs | {} |
| Deals Perdus | {} |

**Périmètre :** {}

---
*Généré automatiquement par le CRM Asset Management — Amundi Edition v15.0*
""".format(date.today().strftime("%d/%m/%Y"), _comex_note,
           _aum_f, _pip_a, _wp, _taux, _nb_f, _nb_a, _nb_l, _perims)

                    st.markdown(md_summary)

                    # Build mailto URL
                    import urllib.parse
                    _subject  = "Rapport Portfolio — {}{}".format(
                        date.today().strftime("%d/%m/%Y"),
                        " [COMEX]" if mode_comex else "")
                    _body  = "Bonjour,\n\nVeuillez trouver ci-dessous la synthese du portfolio :\n\n"
                    _body += "AUM Finance Total : {}\n".format(_aum_f)
                    _body += "Pipeline Actif    : {}\n".format(_pip_a)
                    _body += "Pipeline Pondere  : {}\n".format(_wp)
                    _body += "Taux Conversion   : {}\n".format(_taux)
                    _body += "Deals Funded      : {}\n".format(_nb_f)
                    _body += "Deals Actifs      : {}\n".format(_nb_a)
                    _body += "Deals Perdus      : {}\n\n".format(_nb_l)
                    _body += "Perimetre : {}\n\n".format(_perims)
                    _body += "--\nCRM Asset Management — Amundi Edition v15.0"
                    mailto_url = "mailto:?subject={}&body={}".format(
                        urllib.parse.quote(_subject),
                        urllib.parse.quote(_body))
                    st.link_button("Ouvrir dans Outlook / Gmail",
                                   url=mailto_url,
                                   use_container_width=True)

            except Exception as e:
                st.error("Erreur génération : {}".format(e))

    st.divider()
    st.caption("Version 15.0 — Amundi Edition — Universal Export Hub")


# ---------------------------------------------------------------------------
# EN-TETE + TABS
# ---------------------------------------------------------------------------
st.markdown(
    '<div class="crm-header"><h1>CRM &amp; Reporting — Asset Management</h1>'
    '<p>Pipeline commercial &middot; Suivi des mandats &middot; '
    'Reporting executif &middot; Performance NAV</p></div>',
    unsafe_allow_html=True)

tab_crm, tab_pipeline, tab_dash, tab_sales, tab_activites, tab_settings, tab_perf = st.tabs([
    "CRM Directory", "Pipeline Management", "Executive Dashboard",
    "Sales Tracking", "Activities", "Settings & Admin", "Performance & NAV"])


# ============================================================================
# ONGLET 1 — CRM DIRECTORY  (Split-Screen)
# ============================================================================
with tab_crm:
    st.markdown('<div class="section-title">CRM Directory — Clients &amp; Contacts</div>',
                unsafe_allow_html=True)

    # ── Toolbar ──────────────────────────────────────────────────────────────
    tb1, tb2, tb_spacer = st.columns([1, 1, 6])
    with tb1:
        if st.button("Nouveau compte", key="crm_tb_new_compte", use_container_width=True):
            dialog_add_client()
    with tb2:
        if st.button("Mailing List", key="crm_tb_mailing", use_container_width=True, type="secondary"):
            st.session_state["crm_show_mailing"] = not st.session_state.get("crm_show_mailing", False)

    # ── Mailing List Generator (toggled inline, no expander) ─────────────────
    if st.session_state.get("crm_show_mailing", False):
        st.markdown('<div class="pipeline-hint">Filtrer les contacts pour générer une liste de diffusion ciblée. Seuls les contacts avec un email sont inclus.</div>', unsafe_allow_html=True)
        mg1, mg2, mg3, mg4 = st.columns(4)
        with mg1: ml_regions   = st.multiselect("Region", db.REGIONS_REFERENTIEL, key="ml_regions")
        with mg2: ml_countries = st.multiselect("Country", COUNTRIES_LIST[1:], key="ml_countries")
        with mg3: ml_tiers     = st.multiselect("Tier", db.TIERS_REFERENTIEL, key="ml_tiers")
        with mg4: ml_interests = st.multiselect("Product Interests", db.PRODUCT_INTERESTS, key="ml_interests")
        df_ml = db.get_mailing_list(
            regions=ml_regions or None, countries=ml_countries or None,
            tiers=ml_tiers or None, product_interests=ml_interests or None)
        if df_ml.empty:
            st.info("Aucun contact ne correspond aux filtres.")
        else:
            display_cols = ["first_name","last_name","company","role","email"]
            st.markdown('<div class="pipeline-hint"><b>{}</b> contact(s) trouvé(s)</div>'.format(len(df_ml)),
                        unsafe_allow_html=True)
            st.dataframe(df_ml[display_cols].rename(columns={
                "first_name":"Prénom","last_name":"Nom","company":"Compte",
                "role":"Rôle","email":"Email"}),
                use_container_width=True, hide_index=True,
                height=min(280, 46 + len(df_ml) * 36))
            st.code("; ".join(df_ml["email"].tolist()), language=None)
            st.download_button("Exporter CSV", key="crm_ml_dl",
                data=df_ml[display_cols].to_csv(index=False).encode("utf-8"),
                file_name="mailing_list_{}.csv".format(date.today().isoformat()),
                mime="text/csv")
        st.markdown("---")

    # ── Search ───────────────────────────────────────────────────────────────
    df_hier = db.get_client_hierarchy()
    if df_hier.empty:
        st.markdown(
            '<div style="background:#001c4b04;border:2px dashed #001c4b18;padding:40px;'
            'text-align:center;margin-top:12px;">'
            '<div style="font-size:0.92rem;font-weight:700;color:{m};">Aucun client enregistré</div>'
            '<div style="color:#888;font-size:0.79rem;margin-top:4px;">'
            'Cliquez sur <b>Nouveau compte</b> ci-dessus pour commencer.</div>'
            '</div>'.format(m=MARINE), unsafe_allow_html=True)
    else:
        noms_clients = sorted(df_hier["nom_client"].tolist())
        crm_search   = st.selectbox(
            "Rechercher un compte client…",
            options=noms_clients,
            index=None,
            placeholder="Tapez le nom du client pour rechercher…",
            key="crm_search_box")

        if crm_search is None:
            nb_total   = len(df_hier)
            nb_valide  = int((df_hier["kyc_status"] == "Validé").sum())
            nb_encours = int((df_hier["kyc_status"] == "En cours").sum())
            nb_bloque  = int((df_hier["kyc_status"] == "Bloqué").sum())
            st.markdown(
                '<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin:16px 0;">'
                '<div class="kpi-card kpi-card-static"><div class="kpi-label">Comptes</div>'
                '<div class="kpi-value">{t}</div><div class="kpi-sub">clients enregistrés</div></div>'
                '<div class="kpi-card kpi-card-static"><div class="kpi-label">KYC Validé</div>'
                '<div class="kpi-value" style="color:#22a062;">{v}</div><div class="kpi-sub">&nbsp;</div></div>'
                '<div class="kpi-card kpi-card-static"><div class="kpi-label">KYC En cours</div>'
                '<div class="kpi-value" style="color:{o};">{e}</div><div class="kpi-sub">&nbsp;</div></div>'
                '<div class="kpi-card kpi-card-static"><div class="kpi-label">KYC Bloqué</div>'
                '<div class="kpi-value" style="color:#c0392b;">{b}</div><div class="kpi-sub">&nbsp;</div></div>'
                '</div>'.format(t=nb_total, v=nb_valide, e=nb_encours, b=nb_bloque, o=ORANGE),
                unsafe_allow_html=True)
            st.markdown(
                '<div style="color:#888;font-size:0.78rem;text-align:center;padding:8px 0;">'
                'Sélectionnez un compte dans la liste déroulante pour afficher sa fiche.</div>',
                unsafe_allow_html=True)
        else:
            # ── FICHE CLIENT — SPLIT SCREEN ───────────────────────────────
            sel_row = df_hier[df_hier["nom_client"] == crm_search]
            if sel_row.empty:
                st.warning("Client introuvable.")
            else:
                sel      = sel_row.iloc[0]
                sel_id   = int(sel["id"])
                sel_nom  = str(sel["nom_client"])
                sel_tier = str(sel.get("tier", "Tier 2"))
                sel_kyc  = str(sel.get("kyc_status", "En cours"))
                sel_type = str(sel.get("type_client", ""))
                sel_reg  = str(sel.get("region", ""))
                sel_cty  = str(sel.get("country", ""))
                sel_par  = str(sel.get("parent_nom", ""))
                sel_prod = str(sel.get("product_interests", ""))
                kyc_color = {"Validé": "#22a062", "En cours": ORANGE,
                             "Bloqué": "#c0392b"}.get(sel_kyc, "#aaa")

                # Fetch filiales
                conn_tmp = db.get_connection()
                try:
                    _c = conn_tmp.cursor()
                    _c.execute("SELECT nom_client FROM clients WHERE parent_id = ? ORDER BY nom_client", (sel_id,))
                    filiales = [r[0] for r in _c.fetchall()]
                finally:
                    conn_tmp.close()

                # ── SPLIT SCREEN ──────────────────────────────────────────
                col_gauche, col_droite = st.columns([1.2, 1], gap="large")

                # ════════════════════════════════════════════════════════
                # COLONNE GAUCHE — Entité
                # ════════════════════════════════════════════════════════
                with col_gauche:
                    # Product interest tags
                    prods_html = ""
                    if sel_prod.strip():
                        tags = " ".join(
                            '<span style="background:{c}18;border:1px solid {c}44;color:{m};'
                            'padding:1px 8px;font-size:0.68rem;font-weight:600;margin:2px;">{p}</span>'
                            .format(c=CIEL, m=MARINE, p=p.strip())
                            for p in sel_prod.split(",") if p.strip())
                        prods_html = (
                            '<div style="margin-top:10px;font-size:0.74rem;">'
                            '<span style="color:#888;margin-right:6px;">Product Interests :</span>'
                            + tags + '</div>')

                    st.markdown(
                        '<div class="detail-panel" style="margin-bottom:10px;">'
                        '<div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:10px;">'
                        '<span style="font-size:1.05rem;font-weight:800;color:{m};">{nom}</span>'
                        '{tier_b}'
                        '<span style="background:{kycc};color:#fff;font-size:0.68rem;font-weight:700;padding:2px 10px;">'
                        '{kycd} KYC : {kyc}</span>'
                        '</div>'
                        '<div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;font-size:0.79rem;">'
                        '<div><div style="color:#888;font-size:0.67rem;text-transform:uppercase;letter-spacing:0.5px;">Type</div>'
                        '<div style="font-weight:600;color:{m};">{type}</div></div>'
                        '<div><div style="color:#888;font-size:0.67rem;text-transform:uppercase;letter-spacing:0.5px;">Région</div>'
                        '<div style="font-weight:600;color:{m};">{reg}</div></div>'
                        '<div><div style="color:#888;font-size:0.67rem;text-transform:uppercase;letter-spacing:0.5px;">Tier</div>'
                        '<div style="font-weight:600;color:{m};">{tier}</div></div>'
                        '<div><div style="color:#888;font-size:0.67rem;text-transform:uppercase;letter-spacing:0.5px;">Country</div>'
                        '<div style="font-weight:600;color:{m};">{cty}</div></div>'
                        '</div>'
                        '{prods_html}'
                        '</div>'.format(
                            m=MARINE, nom=sel_nom, tier_b=_tier_badge(sel_tier), tier=sel_tier,
                            kycc=kyc_color, kycd=_kyc_dot(sel_kyc), kyc=sel_kyc,
                            type=sel_type, reg=sel_reg, cty=sel_cty or "—",
                            prods_html=prods_html),
                        unsafe_allow_html=True)

                    _btn_row1, _btn_row2 = st.columns(2)
                    with _btn_row1:
                        if st.button("Modifier le compte",
                                     key="crm_edit_co_{}".format(sel_id),
                                     type="tertiary", use_container_width=True):
                            dialog_edit_client(sel.to_dict())
                    with _btn_row2:
                        if st.button("Meeting Brief",
                                     key="crm_brief_{}".format(sel_id),
                                     type="tertiary", use_container_width=True):
                            _grp = db.get_client_group_summary(sel_id)
                            _ws  = db.get_whitespace_matrix()
                            dialog_meeting_brief(sel.to_dict(), _grp,
                                                 db.get_contacts(sel_id),
                                                 db.get_activities(client_id=sel_id),
                                                 _ws)

                    # PPTX export → Universal Export Hub in sidebar

                    # ── Corporate Structure block ─────────────────────────
                    has_struct = sel_par or filiales
                    if has_struct:
                        st.markdown(
                            '<div style="font-size:0.72rem;font-weight:700;color:{m};'
                            'text-transform:uppercase;letter-spacing:0.6px;'
                            'margin:14px 0 6px 0;border-bottom:1px solid #001c4b15;padding-bottom:3px;">'
                            'Structure du groupe</div>'.format(m=MARINE),
                            unsafe_allow_html=True)
                        if sel_par:
                            st.markdown(
                                '<div class="struct-block">'
                                '<div class="struct-label">Filiale de</div>'
                                '<div class="struct-val">{p}</div>'
                                '</div>'.format(p=sel_par),
                                unsafe_allow_html=True)
                        if filiales:
                            # Enrich filiales with AUM + Tier from group summary
                            _fils_enriched = _grp_summary.get("filiales", [])
                            _fils_dict = {f["nom"]: f for f in _fils_enriched}
                            fils_items = ""
                            for f_nom in filiales:
                                f_data = _fils_dict.get(f_nom, {})
                                f_aum  = fmt_m(f_data.get("aum", 0))
                                f_tier = f_data.get("tier", "—")
                                fils_items += (
                                    '<div style="display:flex;justify-content:space-between;'
                                    'align-items:center;padding:5px 0;'
                                    'border-bottom:1px solid #001c4b08;">'
                                    '<span style="font-weight:600;color:{m};font-size:0.78rem;">{f}</span>'
                                    '<span style="font-size:0.70rem;color:#888;">{tier}</span>'
                                    '<span style="font-size:0.74rem;font-weight:700;color:{ciel};">{aum}</span>'
                                    '</div>').format(m=MARINE, ciel=CIEL,
                                                      f=f_nom, tier=f_tier, aum=f_aum)
                            # Consolidated total
                            aum_grp = fmt_m(_grp_summary.get("aum_consolide", 0))
                            st.markdown(
                                '<div class="struct-block">'
                                '<div style="display:flex;justify-content:space-between;'
                                'margin-bottom:6px;">'
                                '<div class="struct-label">Filiales ({n})</div>'
                                '<div style="font-size:0.72rem;font-weight:700;color:{ciel};">'
                                'Groupe : {aum_grp}</div></div>'
                                '{fils}'
                                '</div>'.format(n=len(filiales), fils=fils_items,
                                              ciel=CIEL, aum_grp=aum_grp),
                                unsafe_allow_html=True)

                    # ── Activités Récentes ────────────────────────────────
                    st.markdown(
                        '<div style="display:flex;align-items:center;justify-content:space-between;'
                        'margin:14px 0 6px 0;border-bottom:1px solid #001c4b15;padding-bottom:3px;">'
                        '<span style="font-size:0.72rem;font-weight:700;color:{m};'
                        'text-transform:uppercase;letter-spacing:0.6px;">Activités récentes</span>'
                        '</div>'.format(m=MARINE),
                        unsafe_allow_html=True)
                    df_act_client = db.get_activities(client_id=sel_id)
                    if df_act_client.empty:
                        st.markdown(
                            '<div style="color:#888;font-size:0.78rem;padding:6px 0;">Aucune activité enregistrée.</div>',
                            unsafe_allow_html=True)
                    else:
                        TYPE_COLORS_CRM = {"Call":CIEL,"Meeting":B_MID,"Email":B_PAL,
                                           "Roadshow":"#2c7fb8","Conference":"#004f8c","Autre":"#888"}
                        for _, act_row in df_act_client.head(3).iterrows():
                            a_typ   = str(act_row.get("type_interaction",""))
                            a_color = TYPE_COLORS_CRM.get(a_typ, "#888")
                            a_notes = str(act_row.get("notes","")) or "—"
                            a_date  = str(act_row.get("date",""))
                            st.markdown(
                                "<div style='border-left:3px solid {c};padding:6px 11px;"
                                "margin:4px 0;background:#f9fbfd;font-size:0.77rem;'>"
                                "<div style='display:flex;justify-content:space-between;'>"
                                "<span style='font-weight:700;color:{c};'>{typ}</span>"
                                "<span style='color:#aaa;font-size:0.68rem;'>{dt}</span></div>"
                                "<div style='color:#444;margin-top:3px;line-height:1.5;'>{notes}</div>"
                                "</div>".format(c=a_color, typ=a_typ, dt=a_date, notes=a_notes[:120]),
                                unsafe_allow_html=True)
                    if st.button("+ Enregistrer une activité", key="crm_add_act_{}".format(sel_id),
                                 type="tertiary"):
                        dialog_add_activity(preselect_client_id=sel_id)

                # ════════════════════════════════════════════════════════
                # COLONNE DROITE — Contacts
                # ════════════════════════════════════════════════════════
                with col_droite:
                    df_contacts = db.get_contacts(sel_id)

                    # Header with inline "Ajouter" button
                    h_left, h_right = st.columns([3, 1])
                    with h_left:
                        st.markdown(
                            '<div style="font-size:0.72rem;font-weight:700;color:{m};'
                            'text-transform:uppercase;letter-spacing:0.6px;">'
                            'Contacts associés ({n})</div>'.format(m=MARINE, n=len(df_contacts)),
                            unsafe_allow_html=True)
                    with h_right:
                        if st.button("+ Ajouter", key="crm_add_ct_{}".format(sel_id),
                                     use_container_width=True):
                            dialog_add_contact(sel_id, sel_nom)

                    if df_contacts.empty:
                        st.markdown(
                            '<div style="color:#888;font-size:0.78rem;padding:10px 0;">'
                            'Aucun contact enregistré.</div>',
                            unsafe_allow_html=True)
                    else:
                        for _, ct in df_contacts.iterrows():
                            prenom   = str(ct.get("prenom", ""))
                            ct_nom   = str(ct.get("nom", ""))
                            role     = str(ct.get("role", ""))
                            email    = str(ct.get("email", "")).strip()
                            tel      = str(ct.get("telephone", "")).strip()
                            linkedin = str(ct.get("linkedin", "")).strip()
                            primary  = bool(ct.get("is_primary", 0))
                            ct_id    = int(ct.get("id", 0))

                            email_html = (
                                '<a href="mailto:{e}" style="color:{ciel};text-decoration:none;">{e}</a>'
                                .format(e=email, ciel=CIEL)) if email else '<span style="color:#bbb;">—</span>'
                            li_url = linkedin
                            if li_url and not li_url.startswith("http"):
                                li_url = "https://" + li_url
                            li_html = (
                                '<a href="{url}" target="_blank" style="color:{ciel};text-decoration:none;">LinkedIn</a>'
                                .format(url=li_url, ciel=CIEL)) if linkedin else ""
                            primary_badge = (
                                '<span style="background:{o};color:#fff;font-size:0.60rem;'
                                'font-weight:700;padding:1px 7px;margin-left:6px;">Principal</span>'
                                .format(o=ORANGE)) if primary else ""

                            # Contact card
                            st.markdown(
                                '<div style="background:#f4f8fc;border-left:3px solid {ciel};'
                                'padding:9px 12px;margin:5px 0;">'
                                '<div style="display:flex;align-items:center;gap:4px;margin-bottom:4px;">'
                                '<span style="font-weight:700;font-size:0.84rem;color:{m};">'
                                '{prenom} {nom}</span>{primary}'
                                '<span style="color:#888;font-size:0.70rem;margin-left:auto;">{role}</span>'
                                '</div>'
                                '<div style="font-size:0.74rem;line-height:1.8;">'
                                '<span style="color:#666;margin-right:12px;">Tél : {tel}</span>'
                                '{email}&nbsp;&nbsp;{li}'
                                '</div></div>'.format(
                                    ciel=CIEL, m=MARINE,
                                    prenom=prenom, nom=ct_nom,
                                    primary=primary_badge, role=role,
                                    tel=tel if tel else "—",
                                    email=email_html, li=li_html),
                                unsafe_allow_html=True)

                            # Action buttons — compact [name+role | Modifier | Supprimer]
                            btn_sp, btn_mod, btn_del = st.columns([8, 1, 1])
                            with btn_mod:
                                if st.button("Modifier", key="crm_edit_ct_{}".format(ct_id),
                                             type="tertiary", help="Modifier ce contact"):
                                    dialog_edit_contact(ct.to_dict())
                            with btn_del:
                                if st.button("Supprimer", key="crm_del_ct_{}".format(ct_id),
                                             type="tertiary", help="Supprimer ce contact"):
                                    st.session_state["crm_del_ct_confirm"] = ct_id

                            if st.session_state.get("crm_del_ct_confirm") == ct_id:
                                st.warning("Supprimer {} {} ?".format(prenom, ct_nom))
                                cdc1, cdc2, _ = st.columns([1, 1, 6])
                                with cdc1:
                                    if st.button("Confirmer", key="crm_del_ct_yes_{}".format(ct_id)):
                                        ok, err = db.delete_contact(ct_id)
                                        if ok:
                                            st.session_state.pop("crm_del_ct_confirm", None)
                                            st.rerun()
                                        else:
                                            st.error(err)
                                with cdc2:
                                    if st.button("Annuler", key="crm_del_ct_no_{}".format(ct_id)):
                                        st.session_state.pop("crm_del_ct_confirm", None)
                                        st.rerun()


# ============================================================================
# ONGLET 2 — PIPELINE MANAGEMENT
# ============================================================================
with tab_pipeline:
    st.markdown('<div class="section-title">Pipeline Management</div>',
                unsafe_allow_html=True)

    # ── Quick Action Center — 4 primary buttons, equal width ────────────────
    qac1, qac2, qac3, qac4 = st.columns(4)
    with qac1:
        if st.button("Nouveau Compte", key="qac_new_compte", use_container_width=True, type="primary"):
            dialog_add_client()
    with qac2:
        if st.button("Nouveau Deal", key="qac_new_deal", use_container_width=True, type="primary"):
            dialog_add_deal()
    with qac3:
        if st.button("Nouvelle Activité", key="qac_new_activity", use_container_width=True, type="primary"):
            dialog_add_activity()
    with qac4:
        if st.button("Nouveau Commercial", key="qac_new_sales", use_container_width=True, type="primary"):
            dialog_manage_sales()
    st.markdown("<br>", unsafe_allow_html=True)

    # ── Filtres ───────────────────────────────────────────────────────────────
    tp_filt, tp_spacer = st.columns([1, 7])
    with tp_filt:
        if st.button("Filtres", key="pipe_filt_toggle", use_container_width=True, type="secondary"):
            st.session_state["pipe_show_filters"] = not st.session_state.get("pipe_show_filters", False)

    if st.session_state.get("pipe_show_filters", False):
        fc1, fc2, fc3 = st.columns(3)
        with fc1: filt_statuts = st.multiselect("Statuts", STATUTS, default=STATUTS_ACTIFS, key="pipe_filt_statuts")
        with fc2: filt_fonds   = st.multiselect("Fonds", FONDS, key="pipe_filt_fonds")
        with fc3: filt_regions = st.multiselect("Regions", REGIONS, key="pipe_filt_regions")
    else:
        filt_statuts = STATUTS_ACTIFS
        filt_fonds   = []
        filt_regions = []

    df_pipe_full = db.get_pipeline_with_last_activity()
    df_view      = df_pipe_full.copy()
    if filt_statuts:  df_view = df_view[df_view["statut"].isin(filt_statuts)]
    if filt_fonds:    df_view = df_view[df_view["fonds"].isin(filt_fonds)]
    if filt_regions:  df_view = df_view[df_view["region"].isin(filt_regions)]

    st.markdown('<div class="pipeline-hint">Sélectionnez une ligne puis cliquez'
                ' "Modifier le deal sélectionné" — <b>{} deal(s)</b> affiché(s)</div>'.format(len(df_view)),
                unsafe_allow_html=True)

    df_display = df_view.copy()
    df_display["next_action_date"] = df_display["next_action_date"].apply(
        lambda d: d.isoformat() if isinstance(d, date) else "")
    # Smart AUM: single column AUM_Pipeline for active deals, funded_aum for Funded
    df_display["aum_pipeline"] = df_display.apply(
        lambda r: float(r["funded_aum"]) if r["statut"] == "Funded"
        else (float(r["revised_aum"]) if float(r["revised_aum"]) > 0
              else float(r["target_aum_initial"])), axis=1)
    df_display["aum_pipeline_fmt"] = df_display["aum_pipeline"].apply(fmt_m)
    df_display["funded_aum_fmt"]   = df_display["funded_aum"].apply(fmt_m)
    if "closing_probability" in df_display.columns:
        df_display["closing_probability"] = df_display["closing_probability"].fillna(50)

    cols_show = ["id","nom_client","type_client","region","fonds","statut",
                 "aum_pipeline_fmt","funded_aum_fmt",
                 "closing_probability","raison_perte","next_action_date","sales_owner","derniere_activite"]

    event = st.dataframe(
        df_display[cols_show], use_container_width=True, height=400, hide_index=True,
        on_select="rerun", selection_mode="single-row",
        column_config={
            "id":                  st.column_config.NumberColumn("ID", width="small"),
            "nom_client":          st.column_config.TextColumn("Client"),
            "type_client":         st.column_config.TextColumn("Type", width="small"),
            "region":              st.column_config.TextColumn("Region", width="small"),
            "fonds":               st.column_config.TextColumn("Fonds"),
            "statut":              st.column_config.TextColumn("Statut"),
            "aum_pipeline_fmt":    st.column_config.TextColumn("AUM Pipeline"),
            "funded_aum_fmt":      st.column_config.TextColumn("AUM Financé"),
            "raison_perte":        st.column_config.TextColumn("Raison"),
            "next_action_date":    st.column_config.TextColumn("Next Action"),
            "sales_owner":         st.column_config.TextColumn("Commercial"),
            "closing_probability": st.column_config.NumberColumn("Proba %", format="%.0f%%", width="small"),
            "derniere_activite":   st.column_config.TextColumn("Dernière Activité"),
        }, key="pipeline_ro")

    # ── Cross-tab safe selection: row index stored locally, dialog only on button click
    selected_rows = event.selection.rows if event.selection else []
    _pipe_selected_id = None
    if selected_rows and selected_rows[0] < len(df_view):
        _pipe_selected_id = int(df_view.iloc[selected_rows[0]]["id"])
        st.session_state["pipe_last_selected_id"] = _pipe_selected_id

    # Show which deal is selected + edit button — dialog ONLY fires on explicit click
    _current_sel = st.session_state.get("pipe_last_selected_id")
    if _current_sel is not None:
        _sel_row_data = df_view[df_view["id"] == _current_sel]
        if not _sel_row_data.empty:
            _sn = str(_sel_row_data.iloc[0].get("nom_client",""))
            _sf = str(_sel_row_data.iloc[0].get("fonds",""))
            _ss = str(_sel_row_data.iloc[0].get("statut",""))
            _hint_col, _btn_col = st.columns([3, 1])
            with _hint_col:
                st.markdown(
                    '<div class="pipeline-hint" style="margin:6px 0;">'
                    'Deal sélectionné : <b>{}</b> &nbsp;·&nbsp; {} &nbsp;·&nbsp; {}</div>'.format(
                        _sn, _sf, _ss),
                    unsafe_allow_html=True)
            with _btn_col:
                if st.button("Modifier le deal sélectionné", key="pipe_open_dialog_btn",
                             type="tertiary", use_container_width=True):
                    _row = db.get_pipeline_row_by_id(_current_sel)
                    if _row:
                        dialog_edit_pipeline(_current_sel, _row)
        else:
            # Deal no longer exists (was deleted) — clear selection
            st.session_state.pop("pipe_last_selected_id", None)
    else:
        st.markdown(
            '<div style="color:#888;font-size:0.78rem;padding:6px 0 2px 0;">'
            'Sélectionnez une ligne pour activer l\'édition.</div>',
            unsafe_allow_html=True)

    st.divider()
    st.markdown("#### AUM Pipeline par Deal Actif")
    df_viz_raw = db.get_pipeline_with_clients()
    # Smart AUM: actifs uniquement, une seule barre = AUM_Pipeline
    df_viz = df_viz_raw[df_viz_raw["statut"].isin(STATUTS_ACTIFS)].copy()
    df_viz["aum_pipeline"] = df_viz["revised_aum"].where(
        df_viz["revised_aum"] > 0, df_viz["target_aum_initial"])
    df_viz = df_viz[df_viz["aum_pipeline"] > 0].copy()
    df_viz["x_label"] = df_viz["nom_client"].str[:16] + " – " + df_viz["fonds"].str[:12]
    df_viz = df_viz.sort_values("aum_pipeline", ascending=False).head(12)
    if not df_viz.empty:
        # Color by statut
        statut_bar_colors = {
            "Soft Commit": B_MID, "Due Diligence": "#004f8c",
            "Initial Pitch": B_PAL, "Prospect": "#9ecae1",
        }
        bar_colors = [statut_bar_colors.get(s, B_MID) for s in df_viz["statut"]]
        fig_viz = go.Figure(go.Bar(
            x=df_viz["x_label"].tolist(),
            y=df_viz["aum_pipeline"].tolist(),
            marker_color=bar_colors,
            marker_line_color=BLANC, marker_line_width=0.5,
            text=[fmt_m(v) for v in df_viz["aum_pipeline"]],
            textposition="outside", textfont_size=9, textfont_color=MARINE,
            hovertemplate="<b>%{x}</b><br>AUM Pipeline : %{text}<br>Statut : %{customdata}<extra></extra>",
            customdata=df_viz["statut"].tolist()))
        fig_viz.update_layout(
            height=340, paper_bgcolor=BLANC, plot_bgcolor=BLANC,
            font_color=MARINE, bargap=0.22,
            xaxis_tickangle=-25, xaxis_showgrid=False,
            yaxis_showgrid=True, yaxis_gridcolor=GRIS,
            margin=dict(l=10, r=10, t=30, b=10))
        st.plotly_chart(fig_viz, use_container_width=True, config={"displayModeBar": False})
        st.caption("AUM Pipeline = AUM Révisé si > 0, sinon AUM Cible. Deals actifs uniquement.")

    df_lp = db.get_pipeline_with_clients()
    df_lp = df_lp[df_lp["statut"].isin(["Lost","Paused"])].copy()
    if not df_lp.empty:
        st.divider()
        st.markdown("#### Lost / Paused Deals")
        df_lp2 = df_lp[["nom_client","fonds","statut","target_aum_initial",
                         "raison_perte","concurrent_choisi","sales_owner"]].copy()
        df_lp2["target_aum_initial"] = df_lp2["target_aum_initial"].apply(fmt_m)
        st.dataframe(df_lp2, use_container_width=True, hide_index=True)


# ============================================================================
# ONGLET 3 — EXECUTIVE DASHBOARD
# ============================================================================
with tab_dash:
    st.markdown('<div class="section-title">Executive Dashboard</div>',
                unsafe_allow_html=True)

    tf_dash = st.selectbox("Timeframe", TIMEFRAMES, key="tf_dash",
                           label_visibility="collapsed")
    cutoff_dash = _timeframe_cutoff(tf_dash)

    kpis = db.get_kpis()
    # pipeline_actif already uses Smart AUM (CASE WHEN revised_aum > 0 ...) in database.py
    if cutoff_dash is not None:
        import pandas as _pd
        _conn = db.get_connection()
        _cutoff_str = cutoff_dash.isoformat()
        _df_p = _pd.read_sql_query(
            "SELECT p.statut, p.funded_aum, p.revised_aum, p.target_aum_initial"
            " FROM pipeline p WHERE DATE(p.updated_at) >= ?",
            _conn, params=(_cutoff_str,))
        _conn.close()
        if not _df_p.empty:
            kpis["total_funded"]    = float(_df_p[_df_p["statut"]=="Funded"]["funded_aum"].sum())
            _df_actif = _df_p[_df_p["statut"].isin(
                ["Prospect","Initial Pitch","Due Diligence","Soft Commit"])].copy()
            # Smart AUM: revised if >0, else target
            _df_actif["_aum_p"] = _df_actif.apply(
                lambda r: float(r["revised_aum"]) if float(r.get("revised_aum",0) or 0) > 0
                          else float(r.get("target_aum_initial",0) or 0), axis=1)
            kpis["pipeline_actif"] = float(_df_actif["_aum_p"].sum())
            kpis["nb_funded"]       = int((_df_p["statut"]=="Funded").sum())
            kpis["nb_lost"]         = int((_df_p["statut"]=="Lost").sum())
            kpis["nb_paused"]       = int((_df_p["statut"]=="Paused").sum())
            kpis["nb_deals_actifs"] = int(_df_p["statut"].isin(
                ["Prospect","Initial Pitch","Due Diligence","Soft Commit"]).sum())
            nb_fl = kpis["nb_funded"] + kpis["nb_lost"]
            kpis["taux_conversion"] = round(kpis["nb_funded"] / nb_fl * 100, 1) if nb_fl > 0 else 0.0
        else:
            for k in ("total_funded","pipeline_actif","nb_funded","nb_lost",
                      "nb_paused","nb_deals_actifs","taux_conversion"):
                kpis[k] = 0

    nb_lost_paused = kpis["nb_lost"] + kpis.get("nb_paused", 0)

    card_lp = (
        '<div class="kpi-card kpi-card-static">'
        '<div class="kpi-label">Lost / Paused Deals</div>'
        '<div class="kpi-value">{n}</div>'
        '<div class="kpi-sub">{n} deals</div>'
        '</div>'
    ).format(n=nb_lost_paused) if nb_lost_paused > 0 else (
        '<div class="kpi-card kpi-card-static">'
        '<div class="kpi-label">Lost / Paused Deals</div>'
        '<div class="kpi-value">0</div>'
        '<div class="kpi-sub">&nbsp;</div>'
        '</div>'
    )
    st.markdown(
        '<div class="kpi-grid">'
        '<div class="kpi-card kpi-card-static">'
        '<div class="kpi-label">Total Funded AUM</div>'
        '<div class="kpi-value">{aum_f}</div>'
        '<div class="kpi-sub">{nb_f} deal(s) funded</div>'
        '</div>'
        '<div class="kpi-card kpi-card-static">'
        '<div class="kpi-label">Active Pipeline</div>'
        '<div class="kpi-value">{aum_p}</div>'
        '<div class="kpi-sub">{nb_p} active deals</div>'
        '</div>'
        '<div class="kpi-card kpi-card-static">'
        '<div class="kpi-label">Weighted Pipeline</div>'
        '<div class="kpi-value">{wp}</div>'
        '<div class="kpi-sub">probability-weighted</div>'
        '</div>'
        '<div class="kpi-card kpi-card-static">'
        '<div class="kpi-label">Conversion Rate</div>'
        '<div class="kpi-value">{taux:.1f}%</div>'
        '<div class="kpi-sub">{nb_f2} funded / {nb_l} lost</div>'
        '</div>'
        '{card_lp}'
        '</div>'.format(
            aum_f=fmt_m(kpis["total_funded"]), nb_f=kpis["nb_funded"],
            aum_p=fmt_m(kpis["pipeline_actif"]), nb_p=kpis["nb_deals_actifs"],
            wp=fmt_m(kpis.get("weighted_pipeline", 0)),
            taux=kpis["taux_conversion"], nb_f2=kpis["nb_funded"], nb_l=kpis["nb_lost"],
            card_lp=card_lp),
        unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    statut_order = [s for s in STATUTS if kpis["statut_repartition"].get(s, 0) > 0]
    if statut_order:
        pills = ""
        for s in statut_order:
            c_hex = STATUT_COLORS.get(s, GRIS)
            count = kpis["statut_repartition"][s]
            pills += (
                '<div class="statut-pill" style="background:{c}16;border:1px solid {c}44;">'
                '<div style="font-size:0.61rem;color:{marine};font-weight:700;text-transform:uppercase;">{s}</div>'
                '<div style="font-size:1.3rem;font-weight:800;color:{c};">{n}</div>'
                '</div>'
            ).format(s=s, c=c_hex, marine=MARINE, n=count)
        st.markdown(
            '<div class="statut-grid" style="grid-template-columns:repeat({n},1fr);">'
            '{pills}</div>'.format(n=len(statut_order), pills=pills),
            unsafe_allow_html=True)

        with st.expander("Détail par statut", expanded=False):
            tabs_statut = st.tabs(["{} ({})".format(s, kpis["statut_repartition"][s])
                                   for s in statut_order])
            for tab_s, s in zip(tabs_statut, statut_order):
                with tab_s:
                    _content_statut(s, _filtre_effectif)

    df_overdue = db.get_overdue_actions()
    if not df_overdue.empty:
        today = date.today()
        alertes_html = ""
        for _, row in df_overdue.iterrows():
            nad = row.get("next_action_date")
            days_late = (today - nad).days if isinstance(nad, date) else 0
            nad_str   = nad.isoformat() if isinstance(nad, date) else "—"
            owner     = str(row.get("sales_owner","")) or ""
            alertes_html += (
                '<div class="alert-overdue">'
                '<b>{client}</b> — {fonds}'
                ' <span style="color:{ciel};font-weight:600;">({statut})</span>'
                ' — Prevue le <b>{nad}</b>'
                ' &nbsp;<span class="badge-retard">+{days}j</span>'
                '{owner_part}'
                '</div>'
            ).format(
                client=row["nom_client"], fonds=row["fonds"], ciel=CIEL,
                statut=row["statut"], nad=nad_str, days=days_late,
                owner_part=" — {}".format(owner) if owner else "")
        st.markdown(alertes_html, unsafe_allow_html=True)
        with st.expander("{} action(s) en retard — détail".format(len(df_overdue)), expanded=False):
            _content_overdue()

    st.divider()

    gcol1, gcol2, gcol3 = st.columns([1, 1, 1.2], gap="medium")
    with gcol1:
        st.markdown("#### Par Type Client")
        abt = kpis.get("aum_by_type", {})
        if abt:
            fig_type = go.Figure(go.Pie(
                labels=list(abt.keys()), values=list(abt.values()), hole=0.52,
                marker_colors=PALETTE[:len(abt)], marker_line_color=BLANC, marker_line_width=2,
                textinfo="percent", textfont_size=10, textfont_color=BLANC))
            fig_type.add_annotation(text=fmt_m(sum(abt.values())),
                                    x=0.5, y=0.55, font_size=11, font_color=MARINE, showarrow=False)
            fig_type.add_annotation(text="Finance",
                                    x=0.5, y=0.42, font_size=8, font_color=GTXT, showarrow=False)
            fig_type.update_layout(height=300, paper_bgcolor=BLANC, plot_bgcolor=BLANC,
                                   font_color=MARINE, showlegend=True,
                                   legend_x=1.02, legend_y=0.5, legend_font_size=9,
                                   margin=dict(l=0, r=80, t=36, b=10))
            st.plotly_chart(fig_type, use_container_width=True, config={"displayModeBar": False})
    with gcol2:
        st.markdown("#### Par Région")
        aum_reg_dash = db.get_aum_by_region()
        if aum_reg_dash:
            fig_reg = go.Figure(go.Pie(
                labels=list(aum_reg_dash.keys()), values=list(aum_reg_dash.values()), hole=0.52,
                marker_colors=PALETTE[:len(aum_reg_dash)], marker_line_color=BLANC, marker_line_width=2,
                textinfo="percent", textfont_size=10, textfont_color=BLANC))
            fig_reg.add_annotation(text=fmt_m(sum(aum_reg_dash.values())),
                                   x=0.5, y=0.55, font_size=11, font_color=MARINE, showarrow=False)
            fig_reg.add_annotation(text="Finance",
                                   x=0.5, y=0.42, font_size=8, font_color=GTXT, showarrow=False)
            fig_reg.update_layout(height=300, paper_bgcolor=BLANC, plot_bgcolor=BLANC,
                                  font_color=MARINE, showlegend=True,
                                  legend_x=1.02, legend_y=0.5, legend_font_size=9,
                                  margin=dict(l=0, r=80, t=36, b=10))
            st.plotly_chart(fig_reg, use_container_width=True, config={"displayModeBar": False})
    with gcol3:
        st.markdown("#### AUM par Fonds")
        abf = kpis.get("aum_by_fonds", {})
        if abf:
            fs = sorted(abf.items(), key=lambda x: x[1], reverse=True)
            fig_fonds = go.Figure(go.Bar(
                x=[v for _, v in fs], y=[k for k, _ in fs], orientation="h",
                marker_color=PALETTE[:len(fs)], marker_line_color=BLANC, marker_line_width=0.5,
                text=[fmt_m(v) for _, v in fs], textposition="outside",
                textfont_size=9, textfont_color=MARINE))
            fig_fonds.update_layout(height=300, paper_bgcolor=BLANC, plot_bgcolor=BLANC,
                                    font_color=MARINE, xaxis_showgrid=True, xaxis_gridcolor=GRIS,
                                    yaxis=dict(autorange="reversed", automargin=True),
                                    margin=dict(l=180, r=20, t=36, b=10))
            st.plotly_chart(fig_fonds, use_container_width=True, config={"displayModeBar": False})

    st.divider()
    _td_col1, _td_col2 = st.columns([3, 1])
    with _td_col1:
        st.markdown("#### Top Deals — Funded AUM")
    with _td_col2:
        _td_region = st.selectbox("Filtrer par Région", ["Toutes"] + REGIONS,
                                   key="dash_top_deals_region",
                                   label_visibility="collapsed")
    df_funded_top = (db.get_pipeline_with_clients()
                     .query("statut == 'Funded'")
                     .sort_values("funded_aum", ascending=False))
    if _td_region != "Toutes":
        df_funded_top = df_funded_top[df_funded_top["region"] == _td_region]
    df_funded_top = df_funded_top.head(10)
    if not df_funded_top.empty:
        max_f = float(df_funded_top["funded_aum"].max())
        for i, (_, row) in enumerate(df_funded_top.iterrows()):
            val   = float(row["funded_aum"])
            pct   = val / max_f * 100 if max_f > 0 else 0
            bar_c = MARINE if i == 0 else B_MID if i < 3 else B_PAL
            st.markdown(
                '<div style="display:flex;align-items:center;gap:10px;'
                'margin:4px 0;padding:7px 12px;border-bottom:1px solid #e8e8e8;">'
                '<div style="font-size:0.72rem;font-weight:700;color:{ciel};min-width:26px;">No.{rank}</div>'
                '<div style="flex:1;">'
                '<div style="font-size:0.80rem;font-weight:700;color:{marine};">{client}</div>'
                '<div style="font-size:0.67rem;color:#777;">{fonds} &middot; {type} &middot; {region}</div>'
                '<div style="background:{gris};height:3px;margin-top:4px;overflow:hidden;">'
                '<div style="background:{barc};width:{pct:.0f}%;height:100%;"></div></div></div>'
                '<div style="text-align:right;min-width:72px;">'
                '<div style="font-size:0.86rem;font-weight:800;color:{marine};">{aum}</div>'
                '</div></div>'.format(
                    rank=i+1, ciel=CIEL, marine=MARINE, gris=GRIS,
                    client=row["nom_client"], fonds=row["fonds"],
                    type=row["type_client"], region=row["region"],
                    barc=bar_c, pct=pct, aum=fmt_m(val)), unsafe_allow_html=True)

    with st.expander("Détail des deals — Funded / Actifs / Perdus", expanded=False):
        dt1, dt2, dt3 = st.tabs(["Funded", "Pipeline Actif", "Lost & Paused"])
        with dt1: _content_funded(_filtre_effectif)
        with dt2: _content_pipeline(_filtre_effectif)
        with dt3: _content_lost(_filtre_effectif)

    # ── DECISION SUPPORT ──────────────────────────────────────────────────
    st.divider()
    ds_col1, ds_col2 = st.columns([1.1, 1], gap="large")

    with ds_col1:
        st.markdown("#### Money in Motion — Projected Inflows")
        df_cf = db.get_expected_cashflows()
        if df_cf.empty or df_cf["aum_pondere"].sum() == 0:
            st.markdown(
                '<div style="background:#001c4b04;border:1px dashed #001c4b20;'
                'padding:18px;text-align:center;">'
                '<div style="color:{m};font-size:0.84rem;">Aucun deal actif avec prochaine action planifiée.</div>'
                '</div>'.format(m=MARINE), unsafe_allow_html=True)
        else:
            all_months = sorted(df_cf["mois"].unique().tolist())
            fig_cf = go.Figure()
            for i, fonds in enumerate(sorted(df_cf["fonds"].unique())):
                df_f = df_cf[df_cf["fonds"] == fonds]
                month_map = dict(zip(df_f["mois"], df_f["aum_pondere"]))
                y_vals = [month_map.get(m, 0) for m in all_months]
                fig_cf.add_trace(go.Bar(
                    name=fonds, x=all_months, y=y_vals,
                    marker_color=PALETTE[i % len(PALETTE)],
                    marker_line_color=BLANC, marker_line_width=0.5,
                    hovertemplate="<b>{}</b><br>%{{x}}<br>%{{customdata}}<extra></extra>".format(fonds),
                    customdata=[fmt_m(v) for v in y_vals]))
            fig_cf.update_layout(
                barmode="stack", height=280,
                paper_bgcolor=BLANC, plot_bgcolor=BLANC, font_color=MARINE,
                legend_bgcolor=BLANC, legend_bordercolor=GRIS, legend_borderwidth=1,
                legend_font_size=9,
                xaxis_showgrid=False, yaxis_showgrid=True, yaxis_gridcolor=GRIS,
                yaxis_tickformat=".2s", yaxis_title="AUM Pondéré (€)",
                margin=dict(l=10, r=10, t=10, b=10))
            st.plotly_chart(fig_cf, use_container_width=True,
                            config={"displayModeBar": False})
            st.caption("AUM pondéré = AUM Pipeline × Probabilité de closing. Répartition par mois de Next Action.")

    with ds_col2:
        st.markdown("#### Whitespace Analysis — Cross-Sell")
        df_ws = db.get_whitespace_matrix()
        if df_ws.empty:
            st.markdown(
                '<div style="background:#001c4b04;border:1px dashed #001c4b20;'
                'padding:18px;text-align:center;">'
                '<div style="color:{m};font-size:0.84rem;">Aucune donnée disponible.</div>'
                '</div>'.format(m=MARINE), unsafe_allow_html=True)
        else:
            # Build HTML table
            fonds_cols = df_ws.columns.tolist()
            tbl = ('<table style="width:100%;border-collapse:collapse;font-size:0.70rem;">'
                   '<thead><tr style="background:{m};color:#fff;">'.format(m=MARINE))
            tbl += '<th style="padding:5px 6px;text-align:left;">Client</th>'
            for f in fonds_cols:
                tbl += '<th style="padding:5px 4px;text-align:center;max-width:60px;word-break:break-word;">{}</th>'.format(f[:8])
            tbl += '</tr></thead><tbody>'
            for i, (client_nm, row) in enumerate(df_ws.iterrows()):
                bg = "#f4f8fc" if i % 2 == 0 else BLANC
                tbl += '<tr style="background:{};">'.format(bg)
                tbl += '<td style="padding:4px 6px;font-weight:600;color:{m};white-space:nowrap;max-width:110px;overflow:hidden;text-overflow:ellipsis;">{c}</td>'.format(
                    m=MARINE, c=str(client_nm)[:16])
                for f in fonds_cols:
                    val = row[f]
                    if pd.isna(val) or val == 0:
                        # Whitespace — opportunity
                        cell = '<td style="text-align:center;padding:3px;background:#fef6f0;">'                                '<span style="color:#f07d00;font-weight:700;font-size:0.65rem;">—</span></td>'
                    else:
                        cell = '<td style="text-align:center;padding:3px;background:#e6f4ea;">'                                '<span style="color:#22a062;font-weight:700;font-size:0.68rem;">{}</span></td>'.format(
                                   fmt_m(val))
                    tbl += cell
                tbl += '</tr>'
            tbl += '</tbody></table>'
            st.markdown(tbl, unsafe_allow_html=True)
            st.caption("Vert = investi (AUM Funded). Orange — = opportunité cross-sell.")




# ============================================================================
# ONGLET 4 — SALES TRACKING
# ============================================================================
with tab_sales:
    st.markdown('<div class="section-title">Sales Tracking</div>',
                unsafe_allow_html=True)
    df_sm = db.get_sales_metrics()
    df_na = db.get_next_actions_by_sales(days_ahead=30)

    if df_sm.empty:
        st.info("Aucune donnée de pipeline disponible.")
    else:
        n_own  = len(df_sm)
        n_cols = min(n_own, 3)
        s_cols = st.columns(n_cols, gap="medium")
        for i, (_, row) in enumerate(df_sm.iterrows()):
            retard_val  = int(row.get("Retards",0))
            retard_html = (
                '<span class="badge-retard">RETARD : {}</span>'.format(retard_val)
                if retard_val > 0
                else '<span style="color:{};font-size:0.74rem;font-weight:600;">A jour</span>'.format(CIEL))
            with s_cols[i % n_cols]:
                st.markdown(
                    '<div class="sales-card">'
                    '<div class="sales-card-name">{name}</div>'
                    '<div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;">'
                    '<div><div class="sales-metric">Total Deals</div><div class="sales-metric-val">{nb}</div></div>'
                    '<div><div class="sales-metric">Funded</div><div class="sales-metric-val">{funded}</div></div>'
                    '<div><div class="sales-metric">AUM Finance</div><div class="sales-metric-acc">{aum}</div></div>'
                    '<div><div class="sales-metric">Pipeline Actif</div><div class="sales-metric-val">{pipe}</div></div>'
                    '<div><div class="sales-metric">Actifs / Perdus</div><div class="sales-metric-val">{actifs} / {perdus}</div></div>'
                    '<div><div class="sales-metric">Alertes</div><div style="margin-top:2px;">{retard}</div></div>'
                    '</div></div><br>'.format(
                        name=row["Commercial"], nb=int(row["Nb_Deals"]), funded=int(row["Funded"]),
                        aum=fmt_m(float(row["AUM_Finance"])), pipe=fmt_m(float(row["Pipeline_Actif"])),
                        actifs=int(row["Actifs"]), perdus=int(row["Perdus"]), retard=retard_html),
                    unsafe_allow_html=True)

        st.divider()
        st.markdown("#### Funded AUM vs Active Pipeline by Sales Rep")
        if df_sm["AUM_Finance"].sum() > 0:
            fig_sales = go.Figure()
            for lbl, col_key, color in [
                ("AUM Finance",    "AUM_Finance",   MARINE),
                ("Pipeline Actif", "Pipeline_Actif", B_MID)]:
                fig_sales.add_trace(go.Bar(name=lbl, x=df_sm["Commercial"].tolist(),
                                           y=df_sm[col_key].tolist(), marker_color=color,
                                           marker_line_color=BLANC, marker_line_width=0.5))
            fig_sales.update_layout(
                height=300, paper_bgcolor=BLANC, plot_bgcolor=BLANC,
                font_color=MARINE, barmode="group", bargap=0.25,
                legend_bgcolor=BLANC, legend_bordercolor=GRIS, legend_borderwidth=1,
                xaxis_showgrid=False, yaxis_showgrid=True, yaxis_gridcolor=GRIS,
                margin=dict(l=10, r=10, t=20, b=10))
            st.plotly_chart(fig_sales, use_container_width=True, config={"displayModeBar": False})

        st.divider()
        st.markdown("#### Strategic Analysis — Fund Breakdown by Market")
        si_mode = st.radio(
            "Périmètre",
            ["Pipeline Actif (AUM Révisé)", "Funded (AUM Financé)"],
            horizontal=True, key="si_mode")
        df_mfb = db.get_market_fonds_breakdown(
            mode="funded" if "Funded" in si_mode else "pipeline")

        if df_mfb.empty or df_mfb["aum"].sum() == 0:
            st.markdown(
                '<div style="background:#001c4b04;border:1px dashed #001c4b20;'
                'padding:24px;text-align:center;">'
                '<div style="color:{m};font-weight:600;font-size:0.85rem;">'
                'Aucune donnée disponible pour ce périmètre</div>'
                '<div style="color:#888;font-size:0.75rem;margin-top:3px;">'
                'Enregistrez des deals et associez-les à des commerciaux pour activer cette analyse.</div>'
                '</div>'.format(m=MARINE), unsafe_allow_html=True)
        else:
            marches = sorted(df_mfb["marche"].unique().tolist())
            fonds_presents = sorted(df_mfb["fonds"].unique().tolist())
            fig_si = go.Figure()
            for i, fonds in enumerate(fonds_presents):
                df_f   = df_mfb[df_mfb["fonds"] == fonds]
                y_vals = []
                for m in marches:
                    row_m = df_f[df_f["marche"] == m]
                    y_vals.append(float(row_m["aum"].iloc[0]) if not row_m.empty else 0.0)
                text_vals = [fmt_m(v) for v in y_vals]
                fig_si.add_trace(go.Bar(
                    name=fonds, x=marches, y=y_vals,
                    text=text_vals, textposition="inside",
                    textfont=dict(size=9, color=BLANC),
                    marker_color=PALETTE[i % len(PALETTE)],
                    marker_line_color=BLANC, marker_line_width=0.5,
                    hovertemplate=(
                        "<b>{f}</b><br>Marché : %{{x}}<br>"
                        "AUM : %{{customdata}}<extra></extra>").format(f=fonds),
                    customdata=text_vals))
            fig_si.update_layout(
                barmode="stack", height=360,
                paper_bgcolor=BLANC, plot_bgcolor=BLANC, font_color=MARINE,
                legend_bgcolor=BLANC, legend_bordercolor=GRIS, legend_borderwidth=1,
                legend_font_size=10, legend_title_text="Fonds",
                xaxis_showgrid=False, xaxis_title="Marché Commercial",
                yaxis_showgrid=True, yaxis_gridcolor=GRIS,
                yaxis_title="AUM (EUR)", yaxis_tickformat=".2s",
                margin=dict(l=10, r=10, t=20, b=10))
            st.plotly_chart(fig_si, use_container_width=True, config={"displayModeBar": False})
            with st.expander("Tableau détaillé Marché × Fonds", expanded=False):
                pivot_si = df_mfb.pivot_table(
                    index="marche", columns="fonds", values="aum",
                    aggfunc="sum", fill_value=0)
                pivot_si["TOTAL"] = pivot_si.sum(axis=1)
                pivot_display = pivot_si.copy()
                for col in pivot_display.columns:
                    pivot_display[col] = pivot_display[col].apply(fmt_m)
                st.dataframe(pivot_display, use_container_width=True)

        st.divider()
        st.markdown("#### Next Actions — 30 Days")
        if df_na.empty:
            st.info("Aucune action planifiée dans les 30 prochains jours.")
        else:
            owners_na    = ["Tous"] + sorted(df_na["sales_owner"].unique().tolist())
            filter_owner = st.selectbox("Filtrer par commercial", owners_na)
            df_nav = df_na if filter_owner == "Tous" else df_na[df_na["sales_owner"] == filter_owner]
            today  = date.today()
            for _, row in df_nav.iterrows():
                nad = row.get("next_action_date")
                if isinstance(nad, date):
                    delta  = (nad - today).days
                    timing = "Dans {}j".format(delta) if delta >= 0 else "RETARD +{}j".format(abs(delta))
                    dot    = CIEL if delta >= 0 else ORANGE
                    nad_s  = nad.isoformat()
                else:
                    timing = "—"; dot = GRIS; nad_s = "—"
                st.markdown(
                    '<div style="display:flex;align-items:center;gap:11px;padding:7px 11px;'
                    'margin:3px 0;background:#f8fafd;border-left:3px solid {dot};">'
                    '<div style="min-width:100px;font-size:0.73rem;color:{dot};font-weight:700;">{timing}</div>'
                    '<div style="flex:1;"><span style="font-weight:600;color:{marine};font-size:0.82rem;">{client}</span>'
                    '&nbsp;<span style="color:#888;font-size:0.73rem;">{fonds} &middot; {statut}</span></div>'
                    '<div style="font-size:0.78rem;color:{ciel};font-weight:600;min-width:68px;text-align:right;">{aum}</div>'
                    '<div style="font-size:0.73rem;color:#888;min-width:85px;text-align:right;">{owner}</div>'
                    '</div>'.format(
                        dot=dot, timing=timing, marine=MARINE, ciel=CIEL,
                        client=row.get("nom_client",""), fonds=row.get("fonds",""),
                        statut=row.get("statut",""), aum=fmt_m(float(row.get("revised_aum",0) or 0) if float(row.get("revised_aum",0) or 0) > 0 else float(row.get("target_aum_initial",0) or 0)),
                        owner=row.get("sales_owner","")), unsafe_allow_html=True)


# ============================================================================
# ONGLET 5 — ACTIVITIES
# ============================================================================
with tab_activites:
    st.markdown('<div class="section-title">Activities Journal</div>',
                unsafe_allow_html=True)

    # Toolbar
    act_tb1, act_tb_spacer = st.columns([1, 7])
    with act_tb1:
        if st.button("Enregistrer une activité", key="act_tb_new_act", use_container_width=True):
            dialog_add_activity()

    tf_act = st.selectbox("Timeframe", TIMEFRAMES, key="tf_act",
                          label_visibility="collapsed")
    cutoff_act = _timeframe_cutoff(tf_act)

    df_all_act = db.get_activities()

    df_act_base = df_all_act.copy()
    if cutoff_act is not None and not df_act_base.empty:
        df_act_base = df_act_base[df_act_base["date"].apply(
            lambda d: (isinstance(d, _date_cls) and d >= cutoff_act)
                      or (isinstance(d, str) and d >= cutoff_act.isoformat())
        )]

    fa1, fa2, fa3 = st.columns([2, 2, 3])
    with fa1:
        clients_for_filter = ["Tous"] + (sorted(df_all_act["nom_client"].unique().tolist()) if not df_all_act.empty else [])
        filt_client = st.selectbox("Client", clients_for_filter, key="filt_act_client")
    with fa2:
        filt_type = st.multiselect("Type d'interaction", TYPES_INTERACTION, key="filt_act_type")
    with fa3:
        filt_search = st.text_input("Recherche dans les notes", placeholder="mot-clé…", key="filt_act_search")

    df_act_filtered = df_all_act.copy() if not df_all_act.empty else df_all_act
    if not df_act_filtered.empty:
        if filt_client != "Tous":
            df_act_filtered = df_act_filtered[df_act_filtered["nom_client"] == filt_client]
        if filt_type:
            df_act_filtered = df_act_filtered[df_act_filtered["type_interaction"].isin(filt_type)]
        if filt_search.strip():
            df_act_filtered = df_act_filtered[
                df_act_filtered["notes"].str.contains(filt_search.strip(), case=False, na=False)]

    if not df_act_base.empty:
        sc1, sc2, sc3, sc4 = st.columns(4)
        nb_mois = len(df_act_base[df_act_base["date"].apply(
            lambda d: str(d)[:7] == date.today().strftime("%Y-%m") if d else False)])
        top_type = df_act_base["type_interaction"].value_counts()
        top_type_str = top_type.index[0] if not top_type.empty else "—"
        for col, label, val in [
            (sc1, "Total Activities",  str(len(df_act_base))),
            (sc2, "This Month",        str(nb_mois)),
            (sc3, "Top Activity Type", top_type_str),
            (sc4, "Clients Reached",   str(df_act_base["nom_client"].nunique())),
        ]:
            with col:
                st.markdown(
                    '<div class="kpi-card kpi-card-static" style="padding:10px;">'
                    '<div class="kpi-label">{}</div>'
                    '<div class="kpi-value" style="font-size:1.15rem;">{}</div>'
                    '</div>'.format(label, val), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.divider()

    if df_act_filtered.empty:
        st.info("Aucune activité trouvée.")
    else:
        nb_total  = len(df_all_act)
        nb_filtre = len(df_act_filtered)
        st.markdown(
            "<div style='font-size:0.78rem;color:#888;margin-bottom:8px;'>"
            "<b>{}</b> activité(s){}".format(
                nb_filtre,
                " sur {} au total".format(nb_total) if nb_filtre < nb_total else ""
            ) + "</div>", unsafe_allow_html=True)

        TYPE_ICONS  = {"Call":"Call","Meeting":"Meeting","Email":"Email",
                       "Roadshow":"Roadshow","Conference":"Conference","Autre":"Autre"}
        TYPE_COLORS = {"Call":CIEL,"Meeting":B_MID,"Email":B_PAL,
                       "Roadshow":"#2c7fb8","Conference":"#004f8c","Autre":"#888"}

        df_sorted = df_act_filtered.copy()
        df_sorted["_date_str"] = df_sorted["date"].apply(lambda d: str(d) if d else "")
        df_sorted = df_sorted.sort_values("_date_str", ascending=False)

        current_date_label = None
        for _, row in df_sorted.iterrows():
            date_str = str(row.get("date",""))
            typ      = str(row.get("type_interaction","Autre"))
            client   = str(row.get("nom_client",""))
            notes    = str(row.get("notes","")) or "—"
            color    = TYPE_COLORS.get(typ, "#888")
            icon     = TYPE_ICONS.get(typ, "")
            act_id   = int(row.get("id", 0))

            if date_str != current_date_label:
                current_date_label = date_str
                try:
                    d_obj  = date.fromisoformat(date_str)
                    delta  = (date.today() - d_obj).days
                    label_d = d_obj.strftime("%A %d %B %Y").capitalize()
                    if   delta == 0: suffix = " — <b style='color:#019ee1;'>Aujourd'hui</b>"
                    elif delta == 1: suffix = " — <span style='color:#888;'>Hier</span>"
                    elif delta <= 7: suffix = " — <span style='color:#888;'>Il y a {} jours</span>".format(delta)
                    else:            suffix = ""
                except Exception:
                    label_d = date_str; suffix = ""
                st.markdown(
                    "<div style='font-size:0.72rem;font-weight:700;color:#7ab8d8;"
                    "text-transform:uppercase;letter-spacing:0.8px;"
                    "padding:10px 0 4px 0;border-bottom:1px solid #e8e8e8;margin-bottom:4px;'>"
                    "{}{}</div>".format(label_d, suffix), unsafe_allow_html=True)

            st.markdown(
                "<div style='display:flex;gap:12px;align-items:flex-start;"
                "padding:9px 14px;margin:3px 0;background:#f9fbfd;"
                "border-left:3px solid {color};'>"
                "<div style='font-size:0.63rem;font-weight:700;color:{color};"
                "background:{color}15;padding:2px 5px;white-space:nowrap;"
                "margin-top:2px;'>{icon}</div>"
                "<div style='flex:1;'>"
                "<div style='font-size:0.82rem;font-weight:700;color:#001c4b;'>"
                "{client}&nbsp;"
                "<span style='font-size:0.69rem;font-weight:600;color:{color};"
                "background:{color}18;padding:1px 8px;'>{typ}</span></div>"
                "<div style='font-size:0.78rem;color:#444;margin-top:4px;line-height:1.55;'>{notes}</div>"
                "</div></div>".format(
                    color=color, icon=icon, client=client, typ=typ, notes=notes),
                unsafe_allow_html=True)

            btn_col1, btn_col2, btn_spacer = st.columns([1, 1, 8])
            with btn_col1:
                if st.button("Modifier", key="act_edit_{}".format(act_id),
                             type="tertiary", help="Modifier cette activité"):
                    dialog_edit_activity(act_id, {
                        "date": date_str, "notes": notes, "type": typ})
            with btn_col2:
                if st.button("Supprimer", key="act_del_{}".format(act_id),
                             type="tertiary", help="Supprimer cette activité"):
                    st.session_state["act_del_confirm"] = act_id

            if st.session_state.get("act_del_confirm") == act_id:
                st.warning("Supprimer cette activité ? Action irréversible.")
                dc1, dc2, _ = st.columns([1, 1, 6])
                with dc1:
                    if st.button("Confirmer", key="act_del_yes_{}".format(act_id)):
                        ok, err = db.delete_activity(act_id)
                        if ok:
                            st.session_state.pop("act_del_confirm", None)
                            st.rerun()
                        else:
                            st.error(err)
                with dc2:
                    if st.button("Annuler", key="act_del_no_{}".format(act_id)):
                        st.session_state.pop("act_del_confirm", None)
                        st.rerun()


# ============================================================================
# ONGLET 6 — SETTINGS & ADMIN
# ============================================================================
with tab_settings:
    st.markdown('<div class="section-title">Settings &amp; Admin</div>',
                unsafe_allow_html=True)

    sa_col1, sa_col2 = st.columns([1, 1], gap="large")

    with sa_col1:
        st.markdown("#### Équipe commerciale")
        st_team_disp = db.get_sales_team()
        if not st_team_disp.empty:
            st.dataframe(st_team_disp[["nom","marche"]], hide_index=True,
                         use_container_width=True,
                         height=min(300, 46 + len(st_team_disp) * 36))
        if st.button("Ajouter un commercial", key="settings_btn_add_sales", use_container_width=True):
            dialog_manage_sales()

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("#### Actions rapides")
        qa1, qa2 = st.columns(2)
        with qa1:
            if st.button("Nouveau compte client", key="settings_btn_new_client", use_container_width=True):
                dialog_add_client()
        with qa2:
            if st.button("Enregistrer une activité", key="settings_btn_new_act", use_container_width=True):
                dialog_add_activity()

    with sa_col2:
        st.markdown("#### Import CSV / Excel — Upsert")
        import_type   = st.radio("Table cible", ["Clients","Pipeline"], horizontal=True)
        uploaded_file = st.file_uploader("Fichier CSV ou Excel (.xlsx)",
                                         type=["csv","xlsx","xls"])
        if import_type == "Clients":
            st.info("Colonnes : nom_client, type_client, region")
        else:
            st.info("Colonnes : nom_client, fonds, statut, target_aum_initial, "
                    "revised_aum, funded_aum, raison_perte, concurrent_choisi, "
                    "next_action_date, sales_owner")
        if uploaded_file:
            try:
                df_imp = (pd.read_csv(uploaded_file) if uploaded_file.name.endswith(".csv")
                          else pd.read_excel(uploaded_file))
                st.dataframe(df_imp.head(5), use_container_width=True, height=145)
                st.caption("{} ligne(s)".format(len(df_imp)))
                if st.button("Lancer l'import", key="settings_btn_import", use_container_width=True):
                    fn = (db.upsert_clients_from_df if import_type == "Clients"
                          else db.upsert_pipeline_from_df)
                    ins, upd = fn(df_imp)
                    st.success("Import : {} créé(s), {} mis à jour.".format(ins, upd))
            except Exception as e:
                st.error("Erreur : {}".format(e))


# ============================================================================
# ONGLET 7 — PERFORMANCE & NAV
# ============================================================================
with tab_perf:
    st.markdown('<div class="section-title">Performance et NAV</div>', unsafe_allow_html=True)

    col_up, col_info = st.columns([1.2, 1], gap="large")

    with col_info:
        st.markdown(
            '<div style="background:#001c4b07;border:1px solid #001c4b14;padding:15px 17px;">'
            '<div style="font-weight:700;color:#001c4b;font-size:0.90rem;margin-bottom:9px;">Format attendu</div>'
            '<div style="font-size:0.79rem;color:#444;line-height:1.75;">'
            '<b>Colonnes obligatoires</b><br>'
            '&bull; <code>Date</code> — YYYY-MM-DD ou DD/MM/YYYY<br>'
            '&bull; <code>Fonds</code> — Nom du fonds<br>'
            '&bull; <code>NAV</code> — Valeur liquidative numerique<br><br>'
            '<b>Calculs produits</b><br>'
            '&bull; Base 100 normalisee &bull; Perf 1M &bull; Perf YTD<br><br>'
            '<b>Export PDF</b><br>'
            'Les donnees NAV sont integrees dans le PDF si cochees.'
            '</div></div>', unsafe_allow_html=True)
        st.markdown("#### Démonstration")
        if st.button("Generer un fichier NAV de demonstration", key="perf_btn_demo_nav", use_container_width=True):
            demo_dates = pd.date_range("{}-01-01".format(date.today().year - 1),
                                       date.today(), freq="B")
            rng  = np.random.default_rng(42)
            navs = {f: 100.0 for f in FONDS}
            rows = []
            for d in demo_dates:
                for fonds, nav in navs.items():
                    nav *= (1 + rng.normal(0.0003, 0.006))
                    navs[fonds] = nav
                    rows.append({"Date": d.date().isoformat(), "Fonds": fonds, "NAV": round(nav, 4)})
            buf_demo = io.BytesIO()
            pd.DataFrame(rows).to_excel(buf_demo, index=False)
            buf_demo.seek(0)
            st.download_button("Telecharger nav_demo.xlsx", data=buf_demo,
                               file_name="nav_demo.xlsx", mime="application/vnd.ms-excel",
                               use_container_width=True)

    with col_up:
        nav_file = st.file_uploader("Charger l'historique NAV (Excel ou CSV)",
                                    type=["xlsx","xls","csv"])

    if nav_file is not None:
        try:
            df_nav = (pd.read_csv(nav_file) if nav_file.name.endswith(".csv")
                      else pd.read_excel(nav_file))
            df_nav.columns = [c.strip() for c in df_nav.columns]
            missing = [c for c in ["Date","Fonds","NAV"] if c not in df_nav.columns]
            if missing:
                st.error("Colonnes manquantes : {}".format(missing)); st.stop()
            df_nav["Date"] = pd.to_datetime(df_nav["Date"], format="mixed", dayfirst=True, errors="coerce")
            df_nav = df_nav.dropna(subset=["Date"])
            df_nav["NAV"]   = pd.to_numeric(df_nav["NAV"], errors="coerce")
            df_nav = df_nav.dropna(subset=["NAV"])
            df_nav["Fonds"] = df_nav["Fonds"].astype(str).str.strip()
            df_nav = df_nav.sort_values("Date").reset_index(drop=True)
            if df_nav.empty:
                st.error("Aucune donnée valide."); st.stop()

            fonds_list = sorted(df_nav["Fonds"].unique().tolist())
            d_min = df_nav["Date"].min(); d_max = df_nav["Date"].max()
            st.markdown(
                '<div style="background:#019ee114;border-left:3px solid #019ee1;'
                'padding:8px 14px;margin:9px 0;">'
                '{:,} points &mdash; {} fonds &mdash; {} au {}</div>'.format(
                    len(df_nav), len(fonds_list),
                    d_min.strftime('%d/%m/%Y'), d_max.strftime('%d/%m/%Y')),
                unsafe_allow_html=True)

            st.markdown("---")
            ff1, ff2, ff3 = st.columns([2, 1, 1])
            with ff1:
                fonds_sel_nav = st.multiselect("Fonds à afficher", fonds_list,
                                               default=fonds_list[:min(5, len(fonds_list))])
            with ff2: d_debut = st.date_input("Depuis", value=d_min.date())
            with ff3: d_fin   = st.date_input("Jusqu'au", value=d_max.date())

            if not fonds_sel_nav:
                st.warning("Sélectionnez au moins un fonds."); st.stop()

            mask = (df_nav["Fonds"].isin(fonds_sel_nav) &
                    (df_nav["Date"].dt.date >= d_debut) &
                    (df_nav["Date"].dt.date <= d_fin))
            df_fn = df_nav[mask].copy()
            if df_fn.empty:
                st.warning("Aucune donnée pour la période."); st.stop()

            pivot = (df_fn.pivot_table(index="Date", columns="Fonds", values="NAV", aggfunc="last")
                     .sort_index().ffill())
            pivot = pivot[[f for f in fonds_sel_nav if f in pivot.columns]]

            base100 = pivot.copy() * np.nan
            for fonds in pivot.columns:
                s = pivot[fonds].dropna()
                if s.empty: continue
                first_val = float(s.iloc[0])
                if first_val != 0 and not np.isnan(first_val):
                    base100[fonds] = pivot[fonds] / first_val * 100

            st.session_state["nav_base100"] = base100
            st.session_state["perf_fonds"]  = fonds_sel_nav

            st.markdown("#### Evolution NAV — Base 100")
            NAV_COLORS = [MARINE, CIEL, B_MID, B_PAL, B_DEP, "#2c7fb8","#004f8c","#6baed6"]
            fig_nav = go.Figure()
            for i, fonds in enumerate(pivot.columns):
                series = base100[fonds].dropna()
                if series.empty: continue
                fig_nav.add_trace(go.Scatter(
                    x=series.index.tolist(), y=series.values.tolist(),
                    mode="lines", name=fonds,
                    line=dict(color=NAV_COLORS[i % len(NAV_COLORS)], width=2),
                    hovertemplate="<b>{}</b><br>%{{y:.2f}}<extra></extra>".format(fonds)))
            fig_nav.add_hline(y=100, line_dash="dot", line_color=GRIS, line_width=1)
            fig_nav.update_layout(
                title_text="Performance NAV — Base 100", title_font_size=13,
                title_font_color=MARINE, height=380,
                paper_bgcolor=BLANC, plot_bgcolor=BLANC, font_color=MARINE,
                legend_bgcolor=BLANC, legend_bordercolor=GRIS, legend_borderwidth=1,
                legend_font_size=10, xaxis_showgrid=False,
                yaxis_showgrid=True, yaxis_gridcolor=GRIS,
                margin=dict(l=10, r=10, t=40, b=10))
            st.plotly_chart(fig_nav, use_container_width=True, config={"displayModeBar": True})

            today_ts  = pd.Timestamp(date.today())
            one_m_ago = today_ts - pd.DateOffset(months=1)
            jan_1     = pd.Timestamp("{}-01-01".format(date.today().year))

            perf_rows = []
            for fonds in pivot.columns:
                series = pivot[fonds].dropna()
                if series.empty: continue
                nav_last = float(series.iloc[-1]); nav_first = float(series.iloc[0])
                s_1m  = series[series.index >= one_m_ago]
                p1m   = ((nav_last / float(s_1m.iloc[0]) - 1) * 100
                         if len(s_1m) > 0 and float(s_1m.iloc[0]) != 0 else float("nan"))
                s_ytd = series[series.index >= jan_1]
                pytd  = ((nav_last / float(s_ytd.iloc[0]) - 1) * 100
                         if len(s_ytd) > 0 and float(s_ytd.iloc[0]) != 0 else float("nan"))
                pp    = ((nav_last / nav_first - 1) * 100 if nav_first != 0 else float("nan"))
                b100s = base100[fonds].dropna()
                nb100 = float(b100s.iloc[-1]) if not b100s.empty else float("nan")
                perf_rows.append({
                    "Fonds": fonds, "NAV Derniere": round(nav_last, 4),
                    "Base 100 Actuel": round(nb100, 2) if not np.isnan(nb100) else None,
                    "Perf 1M (%)":     round(p1m,  2) if not np.isnan(p1m)  else None,
                    "Perf YTD (%)":    round(pytd, 2) if not np.isnan(pytd) else None,
                    "Perf Periode (%)": round(pp, 2) if not np.isnan(pp)    else None})

            if perf_rows:
                df_pt = pd.DataFrame(perf_rows)
                st.session_state["perf_data"] = df_pt
                st.markdown("#### Tableau des Performances")

                def _fp(val):
                    if val is None or (isinstance(val, float) and np.isnan(val)):
                        return '<span style="color:#999;">n.d.</span>'
                    c = CIEL if val >= 0 else "#8b2020"
                    s = "+" if val > 0 else ""
                    return '<span style="color:{};font-weight:700;">{}{:.2f}%</span>'.format(c,s,val)

                tbl_h = (
                    '<table style="width:100%;border-collapse:collapse;font-size:0.82rem;">'
                    '<thead><tr style="background:{marine};color:white;">'
                    '<th style="padding:8px 12px;text-align:left;">Fonds</th>'
                    '<th style="padding:8px 12px;text-align:right;">NAV</th>'
                    '<th style="padding:8px 12px;text-align:right;">Base 100</th>'
                    '<th style="padding:8px 12px;text-align:right;">Perf 1M</th>'
                    '<th style="padding:8px 12px;text-align:right;">Perf YTD</th>'
                    '<th style="padding:8px 12px;text-align:right;">Perf Période</th>'
                    '</tr></thead><tbody>').format(marine=MARINE)
                for i, r in enumerate(perf_rows):
                    bg = "#f0f5fa" if i % 2 == 0 else BLANC
                    tbl_h += (
                        '<tr style="background:{bg};border-bottom:1px solid {gris};">'
                        '<td style="padding:7px 12px;font-weight:600;color:{marine};">{fonds}</td>'
                        '<td style="padding:7px 12px;text-align:right;color:{marine};">{nav}</td>'
                        '<td style="padding:7px 12px;text-align:right;color:{marine};">{b100}</td>'
                        '<td style="padding:7px 12px;text-align:right;">{p1m}</td>'
                        '<td style="padding:7px 12px;text-align:right;">{pytd}</td>'
                        '<td style="padding:7px 12px;text-align:right;">{pp}</td>'
                        '</tr>').format(
                        bg=bg, gris=GRIS, marine=MARINE, fonds=r["Fonds"],
                        nav="{:.4f}".format(r["NAV Derniere"]),
                        b100="{:.2f}".format(r["Base 100 Actuel"]) if r["Base 100 Actuel"] else "n.d.",
                        p1m=_fp(r["Perf 1M (%)"]), pytd=_fp(r["Perf YTD (%)"]),
                        pp=_fp(r["Perf Periode (%)"]))
                tbl_h += "</tbody></table>"
                st.markdown(tbl_h, unsafe_allow_html=True)
                st.markdown("<br>", unsafe_allow_html=True)
                st.download_button("Exporter le tableau en CSV",
                                   data=df_pt.to_csv(index=False).encode("utf-8"),
                                   file_name="performances_{}.csv".format(date.today().isoformat()),
                                   mime="text/csv")

                ytd_data = [r for r in perf_rows if r["Perf YTD (%)"] is not None]
                if ytd_data:
                    ytd_data.sort(key=lambda r: r["Perf YTD (%)"], reverse=True)
                    st.markdown("#### Comparaison Performances YTD")
                    fig_ytd = go.Figure(go.Bar(
                        x=[r["Fonds"] for r in ytd_data],
                        y=[r["Perf YTD (%)"] for r in ytd_data],
                        marker_color=[CIEL if v >= 0 else GRIS
                                      for v in [r["Perf YTD (%)"] for r in ytd_data]],
                        marker_line_color=BLANC, marker_line_width=0.4,
                        text=["{:+.2f}%".format(r["Perf YTD (%)"]) for r in ytd_data],
                        textposition="outside", textfont_size=9, textfont_color=MARINE))
                    fig_ytd.add_hline(y=0, line_color=MARINE, line_width=0.8)
                    fig_ytd.update_layout(height=300, paper_bgcolor=BLANC, plot_bgcolor=BLANC,
                                          font_color=MARINE, xaxis_tickangle=-15,
                                          xaxis_showgrid=False, yaxis_showgrid=True,
                                          yaxis_gridcolor=GRIS, margin=dict(l=10,r=10,t=36,b=10))
                    st.plotly_chart(fig_ytd, use_container_width=True,
                                    config={"displayModeBar": False})

        except Exception as e:
            st.error("Erreur traitement NAV : {}".format(e))
            import traceback
            with st.expander("Détails de l'erreur"):
                st.code(traceback.format_exc())
    else:
        st.markdown(
            '<div style="background:#001c4b04;border:2px dashed #001c4b1e;'
            'padding:44px;text-align:center;margin-top:12px;">'
            '<div style="font-size:0.96rem;font-weight:700;color:#001c4b;margin-bottom:5px;">'
            'Module Performance et NAV</div>'
            '<div style="color:#777;font-size:0.81rem;max-width:380px;margin:0 auto;line-height:1.65;">'
            'Chargez un fichier Excel ou CSV avec les colonnes '
            '<code>Date</code>, <code>Fonds</code>, <code>NAV</code>.<br><br>'
            'Utilisez le bouton <b>Generer un fichier NAV de demonstration</b>.'
            '</div></div>', unsafe_allow_html=True)
