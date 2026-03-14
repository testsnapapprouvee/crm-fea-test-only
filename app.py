# =============================================================================
# app.py — CRM Asset Management — Amundi Edition
# Charte : #001c4b Marine | #019ee1 Ciel | #f07d00 Orange
# UX : zéro bouton parasite dans le DOM.
#      Technique : st.query_params — les éléments visuels sont de purs liens
#      <a href="?open=X"> sans aucun st.button pour les zones cliquables.
#      Au rechargement Streamlit lit ?open= et déclenche le bon @st.dialog.
# =============================================================================

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from datetime import date, timedelta
import io
import sys
import os

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

STATUT_COLORS = {
    "Funded":        B_MID,    "Soft Commit":   "#2c7fb8",
    "Due Diligence": "#004f8c", "Initial Pitch": B_PAL,
    "Prospect":      "#9ecae1", "Lost":          "#aaaaaa",
    "Paused":        "#c0c0c0", "Redeemed":      "#b8b8d0",
}

fmt_m = db.format_finance

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
st.set_page_config(page_title="CRM - Asset Management",
                   layout="wide", initial_sidebar_state="expanded")

# ---------------------------------------------------------------------------
# CSS — éléments cliquables = liens HTML purs, zéro forme Streamlit
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

/* ============================================================
   ÉLÉMENTS CLIQUABLES — liens HTML purs, zéro bouton Streamlit
   ============================================================ */

/* Lien wrapper commun — retire toute décoration de lien */
a.click-wrap {
    display: block;
    text-decoration: none;
    color: inherit;
    cursor: pointer;
}
a.click-wrap:hover { text-decoration: none; }

/* KPI Cards */
.kpi-card {
    background: #001c4b;
    padding: 14px 10px 12px 10px;
    text-align: center;
    border: 1px solid #1a5e8a;
    border-bottom: 2px solid #019ee1;
    transition: background 0.12s, border-color 0.12s;
    display: block;
}
a.click-wrap:hover .kpi-card {
    background: #0a2d5e;
    border-color: #f07d00 !important;
}
.kpi-label { font-size:0.64rem; color:#7ab8d8; text-transform:uppercase; letter-spacing:0.8px; margin-bottom:5px; font-weight:600; }
.kpi-value { font-size:1.4rem; font-weight:800; color:#ffffff; }
.kpi-sub   { font-size:0.61rem; color:#c8dde8; margin-top:3px; }
.kpi-card-static { border-bottom:2px solid #1a5e8a; }

/* Pastilles statut */
.statut-pill {
    padding: 8px 7px;
    text-align: center;
    transition: filter 0.10s;
    display: block;
}
a.click-wrap:hover .statut-pill { filter: brightness(1.14); }

/* Alertes retard */
.alert-overdue {
    background: #fef6f0;
    border-left: 3px solid #f07d00;
    padding: 8px 12px;
    margin: 3px 0;
    font-size: 0.77rem;
    color: #001c4b;
    transition: background 0.10s;
    display: block;
}
a.click-wrap:hover .alert-overdue { background: #fdebd5; }

/* Activités */
.activity-row {
    border-left: 2px solid #019ee1;
    padding: 7px 11px;
    margin: 3px 0;
    background: #f9fbfd;
    font-size: 0.78rem;
    transition: background 0.10s;
    display: block;
}
a.click-wrap:hover .activity-row { background: #e4f1f9; }

/* ============================================================
   AUTRES ÉLÉMENTS
   ============================================================ */
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

/* Boutons normaux (formulaires, sidebar, PDF) */
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
# HELPER — lien cliquable sans st.button
# Génère <a href="?open=KEY" target="_self" class="click-wrap">HTML</a>
# Streamlit relit les query_params au rechargement → déclenche le dialog.
# ---------------------------------------------------------------------------
def clickable(html_content, open_key):
    """Rend html_content entièrement cliquable via un lien query_param."""
    st.markdown(
        '<a href="?open={key}" target="_self" class="click-wrap">'
        '{content}'
        '</a>'.format(key=open_key, content=html_content),
        unsafe_allow_html=True
    )


def statut_badge(statut):
    colors = {
        "Funded":        (B_MID,       BLANC),  "Soft Commit":   ("#2c7fb8", BLANC),
        "Due Diligence": ("#004f8c",   BLANC),  "Initial Pitch": (B_PAL,    BLANC),
        "Prospect":      ("#001c4b18", MARINE), "Lost":          (GRIS,     "#555"),
        "Paused":        ("#d0d0d0",   MARINE), "Redeemed":      ("#c8d8e8", MARINE),
    }
    bg, fg = colors.get(statut, (GRIS, "#555"))
    return '<span style="padding:2px 9px;font-size:0.71rem;font-weight:600;background:{};color:{};">{}</span>'.format(bg, fg, statut)


# ---------------------------------------------------------------------------
# MODALES @st.dialog
# ---------------------------------------------------------------------------

@st.dialog("Deals Finances (Funded)")
def modal_funded(fonds_filter=None):
    df = db.get_funded_deals_detail(fonds_filter)
    if df.empty:
        st.info("Aucun deal Funded."); return
    df2 = df.copy()
    for col in ["AUM_Finance","AUM_Cible"]:
        if col in df2.columns: df2[col] = df2[col].apply(fmt_m)
    st.dataframe(df2, use_container_width=True, hide_index=True,
                 height=min(420, 46 + len(df2) * 36))


@st.dialog("Pipeline Actif")
def modal_pipeline_actif(fonds_filter=None):
    df = db.get_active_deals_detail(fonds_filter)
    if df.empty:
        st.info("Aucun deal actif."); return
    df2 = df.copy()
    for col in ["AUM_Revise","AUM_Cible"]:
        if col in df2.columns: df2[col] = df2[col].apply(fmt_m)
    if "Prochaine_Action" in df2.columns:
        df2["Prochaine_Action"] = df2["Prochaine_Action"].apply(
            lambda d: d.isoformat() if isinstance(d, date) else str(d or "—"))
    st.dataframe(df2, use_container_width=True, hide_index=True,
                 height=min(420, 46 + len(df2) * 36))


@st.dialog("Deals Perdus et En Pause")
def modal_lost(fonds_filter=None):
    df = db.get_lost_deals_detail(fonds_filter)
    if df.empty:
        st.info("Aucun deal Lost/Paused."); return
    df2 = df.copy()
    if "AUM_Cible" in df2.columns: df2["AUM_Cible"] = df2["AUM_Cible"].apply(fmt_m)
    st.dataframe(df2, use_container_width=True, hide_index=True,
                 height=min(420, 46 + len(df2) * 36))


@st.dialog("Activites et Notes")
def modal_activities_tab(client_nom=None):
    df = db.get_activities()
    if not df.empty and client_nom:
        df = df[df["nom_client"] == client_nom]
    if df.empty:
        st.info("Aucune activite enregistree."); return
    for _, row in df.iterrows():
        st.markdown(
            "<div style='border-left:2px solid #019ee1;padding:6px 12px;"
            "margin:4px 0;background:#f4f8fc;'>"
            "<div style='font-size:0.78rem;font-weight:700;color:#001c4b;'>"
            "{client} &nbsp;<span style='color:#019ee1;'>{type}</span>"
            "&nbsp;<span style='color:#888;font-size:0.70rem;'>{date}</span></div>"
            "<div style='font-size:0.80rem;color:#444;margin-top:3px;'>{notes}</div>"
            "</div>".format(
                client=row.get("nom_client",""), type=row.get("type_interaction",""),
                date=str(row.get("date","")),
                notes=str(row.get("notes","")) or "<em style='color:#aaa;'>Aucune note</em>",
            ), unsafe_allow_html=True)


@st.dialog("Actions en Retard — Detail complet")
def modal_overdue_detail():
    df = db.get_overdue_actions()
    if df.empty:
        st.info("Aucune action en retard."); return
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


@st.dialog("Deals — Statut selectionne")
def modal_statut_detail(statut_nom, fonds_filter=None):
    df = db.get_pipeline_by_statut(statut_nom, fonds_filter)
    if df.empty:
        st.info("Aucun deal en statut {}.".format(statut_nom)); return
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
                aum=fmt_m(float(row.get("revised_aum",0) or 0)),
                nad=nad_str, timing=timing_html, owner=row.get("sales_owner",""),
                act_block=(
                    "<div style='margin-top:5px;font-size:0.73rem;color:#666;"
                    "border-top:1px solid #e8e8e8;padding-top:4px;'>"
                    "<b>Derniere activite :</b> {}</div>".format(act)
                ) if act else ""), unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# LECTURE DU QUERY PARAM — déclenche le bon dialog en haut de page
# (avant tout rendu, pour que le dialog s'ouvre immédiatement)
# ---------------------------------------------------------------------------
_open = st.query_params.get("open", "")
_filtre_effectif = None   # sera recalculé dans la sidebar, mais on en a besoin ici aussi
                           # pour les modals déclenchés par query_param

if _open == "funded":
    modal_funded(None)
    st.query_params.clear()
elif _open == "pipeline":
    modal_pipeline_actif(None)
    st.query_params.clear()
elif _open == "lost":
    modal_lost(None)
    st.query_params.clear()
elif _open == "overdue":
    modal_overdue_detail()
    st.query_params.clear()
elif _open.startswith("statut_"):
    statut_nom = _open[len("statut_"):]
    modal_statut_detail(statut_nom, None)
    st.query_params.clear()
elif _open.startswith("act_"):
    client_nom = _open[len("act_"):]
    modal_activities_tab(client_nom if client_nom else None)
    st.query_params.clear()


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
    st.markdown('<div style="color:#4a8fbd;font-size:0.67rem;font-weight:700;'
                'text-transform:uppercase;letter-spacing:1px;margin-bottom:5px;">Export PDF</div>',
                unsafe_allow_html=True)

    fonds_perimetre = st.multiselect("Perimetre de l'export", options=FONDS, default=FONDS,
                                     key="fonds_perimetre_select")
    _filtre_effectif = (fonds_perimetre
                        if (fonds_perimetre and len(fonds_perimetre) < len(FONDS))
                        else None)

    mode_comex = st.toggle("Mode Comex — Anonymisation", value=False)

    st.markdown('<div style="color:#8ab8d8;font-size:0.64rem;font-weight:600;'
                'text-transform:uppercase;letter-spacing:0.8px;margin:8px 0 4px 0;">Sections</div>',
                unsafe_allow_html=True)
    include_top10    = st.checkbox("Top 10 Inflows", value=True)
    include_outflows = st.checkbox("Top 10 Outflows (Redeemed)", value=False,
                                   disabled=not include_top10)
    include_perf_pdf = st.checkbox("Performance NAV", value=True)

    if fonds_perimetre and len(fonds_perimetre) < len(FONDS):
        st.markdown("<div>{}</div>".format(" ".join(
            '<span class="perimetre-badge">{}</span>'.format(f) for f in fonds_perimetre)),
            unsafe_allow_html=True)
    elif not fonds_perimetre:
        st.warning("Selectionnez au moins un fonds.")

    perf_data_pdf   = st.session_state.get("perf_data",   None)
    nav_base100_pdf = st.session_state.get("nav_base100", None)
    if perf_data_pdf is not None:
        st.caption("NAV chargee : {} fonds.".format(len(perf_data_pdf)))

    if not fonds_perimetre:
        st.button("Generer le rapport PDF", disabled=True, use_container_width=True)
    elif st.button("Generer le rapport PDF", use_container_width=True):
        with st.spinner("Generation en cours..."):
            try:
                pipeline_pdf   = db.get_pipeline_with_clients(fonds_filter=_filtre_effectif)
                kpis_pdf       = db.get_kpis(fonds_filter=_filtre_effectif)
                aum_region_pdf = db.get_aum_by_region(fonds_filter=_filtre_effectif)
                pf_pdf = perf_data_pdf
                nb_pdf = nav_base100_pdf
                if pf_pdf is not None and _filtre_effectif and "Fonds" in pf_pdf.columns:
                    pf_pdf = pf_pdf[pf_pdf["Fonds"].isin(_filtre_effectif)]
                if nb_pdf is not None and _filtre_effectif and hasattr(nb_pdf,"columns"):
                    cols_k = [c for c in nb_pdf.columns if c in _filtre_effectif]
                    nb_pdf = nb_pdf[cols_k] if cols_k else None
                _include_perf = (include_perf_pdf and pf_pdf is not None
                                 and hasattr(pf_pdf,"empty") and not pf_pdf.empty)
                pdf_bytes = pdf_gen.generate_pdf(
                    pipeline_df=pipeline_pdf, kpis=kpis_pdf,
                    aum_by_region=aum_region_pdf, mode_comex=mode_comex,
                    perf_data=pf_pdf, nav_base100_df=nb_pdf,
                    fonds_perimetre=fonds_perimetre,
                    include_top10=include_top10, include_outflows=include_outflows,
                    include_perf=_include_perf)
                fname = "report{}_{}.pdf".format(
                    "_comex" if mode_comex else "", date.today().isoformat())
                st.download_button("Telecharger le rapport", data=pdf_bytes,
                                   file_name=fname, mime="application/pdf",
                                   use_container_width=True)
                st.success("Rapport genere.")
            except Exception as e:
                st.error("Erreur : {}".format(e))

    st.divider()
    st.caption("Version 9.0 — Amundi Edition")


# ---------------------------------------------------------------------------
# EN-TETE
# ---------------------------------------------------------------------------
st.markdown(
    '<div class="crm-header"><h1>CRM &amp; Reporting — Asset Management</h1>'
    '<p>Pipeline commercial &middot; Suivi des mandats &middot; '
    'Reporting executif &middot; Performance NAV</p></div>',
    unsafe_allow_html=True)

tab_ingest, tab_pipeline, tab_dash, tab_sales, tab_perf = st.tabs([
    "Data Ingestion", "Pipeline Management", "Executive Dashboard",
    "Sales Tracking", "Performance & NAV"])


# ============================================================================
# ONGLET 1 — DATA INGESTION
# ============================================================================
with tab_ingest:
    st.markdown('<div class="section-title">Saisie et Import de Donnees</div>',
                unsafe_allow_html=True)
    col_form, col_import = st.columns([1, 1], gap="large")

    with col_form:
        with st.expander("Ajouter un Client", expanded=True):
            with st.form("form_client", clear_on_submit=True):
                c1, c2 = st.columns(2)
                with c1:
                    nom_client  = st.text_input("Nom du Client")
                    type_client = st.selectbox("Type Client", TYPES_CLIENT)
                with c2:
                    region = st.selectbox("Region", REGIONS)
                sub_c = st.form_submit_button("Enregistrer", use_container_width=True)
            if sub_c:
                if not nom_client.strip():
                    st.error("Nom du client obligatoire.")
                else:
                    try:
                        db.add_client(nom_client.strip(), type_client, region)
                        st.success("Client {} ajoute.".format(nom_client))
                    except Exception as e:
                        st.warning("Ce client existe deja." if "UNIQUE" in str(e)
                                   else "Erreur : {}".format(e))

        with st.expander("Ajouter un Deal Pipeline", expanded=True):
            clients_dict = db.get_client_options()
            sales_owners = db.get_sales_owners() or ["Non assigne"]
            if not clients_dict:
                st.info("Ajoutez d'abord un client.")
            else:
                with st.form("form_deal", clear_on_submit=True):
                    ca, cb = st.columns(2)
                    with ca:
                        client_sel  = st.selectbox("Client", list(clients_dict.keys()))
                        fonds_sel   = st.selectbox("Fonds", FONDS)
                        statut_sel  = st.selectbox("Statut", STATUTS)
                        owner_input = st.text_input("Commercial",
                                                    value=sales_owners[0] if sales_owners else "")
                    with cb:
                        target_aum  = st.number_input("AUM Cible (EUR)", min_value=0, step=1_000_000)
                        revised_aum = st.number_input("AUM Revise (EUR)", min_value=0, step=1_000_000)
                        funded_aum  = st.number_input("AUM Finance (EUR)", min_value=0, step=1_000_000)
                    raison_perte, concurrent = "", ""
                    if statut_sel in ("Lost","Paused"):
                        cc, cd = st.columns(2)
                        with cc: raison_perte = st.selectbox("Raison", RAISONS_PERTE)
                        with cd: concurrent   = st.text_input("Concurrent")
                    next_action = st.date_input("Prochaine Action",
                                               value=date.today() + timedelta(days=14))
                    sub_d = st.form_submit_button("Enregistrer le Deal", use_container_width=True)
                if sub_d:
                    if statut_sel in ("Lost","Paused") and not raison_perte:
                        st.error("Raison obligatoire.")
                    else:
                        db.add_pipeline_entry(
                            clients_dict[client_sel], fonds_sel, statut_sel,
                            float(target_aum), float(revised_aum), float(funded_aum),
                            raison_perte, concurrent, next_action.isoformat(),
                            owner_input.strip() or "Non assigne")
                        st.success("Deal {} / {} enregistre.".format(fonds_sel, client_sel))

        with st.expander("Enregistrer une Activite"):
            clients_dict2 = db.get_client_options()
            if clients_dict2:
                with st.form("form_act", clear_on_submit=True):
                    ce, cf = st.columns(2)
                    with ce:
                        act_client = st.selectbox("Client", list(clients_dict2.keys()))
                        act_type   = st.selectbox("Type", TYPES_INTERACTION)
                    with cf:
                        act_date  = st.date_input("Date", value=date.today())
                        act_notes = st.text_area("Notes", height=68)
                    sub_a = st.form_submit_button("Enregistrer", use_container_width=True)
                if sub_a:
                    db.add_activity(clients_dict2[act_client], act_date.isoformat(),
                                    act_notes, act_type)
                    st.success("Activite enregistree pour {}.".format(act_client))

    with col_import:
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
                if st.button("Lancer l'import", use_container_width=True):
                    fn = (db.upsert_clients_from_df if import_type == "Clients"
                          else db.upsert_pipeline_from_df)
                    ins, upd = fn(df_imp)
                    st.success("Import : {} cree(s), {} mis a jour.".format(ins, upd))
            except Exception as e:
                st.error("Erreur : {}".format(e))

        st.divider()
        st.markdown("#### Dernieres Activites")
        df_act = db.get_activities()
        if not df_act.empty:
            for _, arow in df_act.head(8).iterrows():
                notes_full    = str(arow.get("notes",""))
                notes_preview = notes_full[:90] + ("…" if len(notes_full) > 90 else "")
                client_enc    = str(arow.get("nom_client","")).replace(" ", "+")
                # Lien pur — zéro bouton Streamlit
                clickable(
                    "<div class='activity-row'>"
                    "<b>{client}</b> &nbsp;<span style='color:#019ee1;'>{type}</span>"
                    " &nbsp;<span style='color:#888;font-size:0.70rem;'>{date}</span><br/>"
                    "<span style='color:#444;'>{notes}</span>"
                    "</div>".format(
                        client=arow.get("nom_client",""),
                        type=arow.get("type_interaction",""),
                        date=str(arow.get("date","")),
                        notes=notes_preview),
                    open_key="act_{}".format(client_enc)
                )


# ============================================================================
# ONGLET 2 — PIPELINE MANAGEMENT
# ============================================================================
with tab_pipeline:
    st.markdown('<div class="section-title">Pipeline Management</div>',
                unsafe_allow_html=True)

    with st.expander("Filtres", expanded=False):
        fc1, fc2, fc3 = st.columns(3)
        with fc1: filt_statuts = st.multiselect("Statuts", STATUTS, default=STATUTS_ACTIFS)
        with fc2: filt_fonds   = st.multiselect("Fonds", FONDS)
        with fc3: filt_regions = st.multiselect("Regions", REGIONS)

    df_pipe_full = db.get_pipeline_with_last_activity()
    df_view      = df_pipe_full.copy()
    if filt_statuts:  df_view = df_view[df_view["statut"].isin(filt_statuts)]
    if filt_fonds:    df_view = df_view[df_view["fonds"].isin(filt_fonds)]
    if filt_regions:  df_view = df_view[df_view["region"].isin(filt_regions)]

    st.markdown('<div class="pipeline-hint">Selectionnez une ligne pour modifier'
                ' — <b>{} deal(s)</b> affiches</div>'.format(len(df_view)),
                unsafe_allow_html=True)

    df_display = df_view.copy()
    df_display["next_action_date"] = df_display["next_action_date"].apply(
        lambda d: d.isoformat() if isinstance(d, date) else "")
    for col in ["target_aum_initial","revised_aum","funded_aum"]:
        if col in df_display.columns:
            df_display[col] = df_display[col].apply(fmt_m)

    cols_show = ["id","nom_client","type_client","region","fonds","statut",
                 "target_aum_initial","revised_aum","funded_aum",
                 "raison_perte","next_action_date","sales_owner","derniere_activite"]

    event = st.dataframe(
        df_display[cols_show], use_container_width=True, height=400, hide_index=True,
        on_select="rerun", selection_mode="single-row",
        column_config={
            "id":                 st.column_config.NumberColumn("ID", width="small"),
            "nom_client":         st.column_config.TextColumn("Client"),
            "type_client":        st.column_config.TextColumn("Type", width="small"),
            "region":             st.column_config.TextColumn("Region", width="small"),
            "fonds":              st.column_config.TextColumn("Fonds"),
            "statut":             st.column_config.TextColumn("Statut"),
            "target_aum_initial": st.column_config.TextColumn("AUM Cible"),
            "revised_aum":        st.column_config.TextColumn("AUM Revise"),
            "funded_aum":         st.column_config.TextColumn("AUM Finance"),
            "raison_perte":       st.column_config.TextColumn("Raison"),
            "next_action_date":   st.column_config.TextColumn("Next Action"),
            "sales_owner":        st.column_config.TextColumn("Commercial"),
            "derniere_activite":  st.column_config.TextColumn("Derniere Activite"),
        }, key="pipeline_ro")

    selected_rows = event.selection.rows if event.selection else []
    pipeline_id   = None
    if len(selected_rows) > 0 and selected_rows[0] < len(df_view):
        pipeline_id = int(df_view.iloc[selected_rows[0]]["id"])

    if pipeline_id is not None:
        row_data = db.get_pipeline_row_by_id(pipeline_id)
        if row_data:
            client_name    = str(row_data.get("nom_client",""))
            current_statut = str(row_data.get("statut","Prospect"))
            st.markdown(
                '<div class="detail-panel"><div style="font-size:0.88rem;font-weight:700;color:{marine};">'
                'Modification — <span style="color:{ciel};">{name}</span>'
                '&nbsp;{badge}&nbsp;<span style="font-size:0.68rem;color:#888;font-weight:400;">'
                'ID #{pid}</span></div></div>'.format(
                    marine=MARINE, ciel=B_MID, name=client_name,
                    badge=statut_badge(current_statut), pid=pipeline_id),
                unsafe_allow_html=True)
            with st.container():
                with st.form(key="edit_{}".format(pipeline_id)):
                    r1c1, r1c2 = st.columns(2)
                    with r1c1:
                        fi = FONDS.index(row_data["fonds"]) if row_data["fonds"] in FONDS else 0
                        new_fonds = st.selectbox("Fonds", FONDS, index=fi)
                    with r1c2:
                        si = STATUTS.index(current_statut) if current_statut in STATUTS else 0
                        new_statut = st.selectbox("Statut", STATUTS, index=si)
                    r2c1, r2c2, r2c3 = st.columns(3)
                    with r2c1:
                        new_target = st.number_input("AUM Cible (EUR)",
                            value=float(row_data.get("target_aum_initial",0.0)),
                            min_value=0.0, step=1_000_000.0, format="%.0f")
                    with r2c2:
                        new_revised = st.number_input("AUM Revise (EUR)",
                            value=float(row_data.get("revised_aum",0.0)),
                            min_value=0.0, step=1_000_000.0, format="%.0f")
                    with r2c3:
                        new_funded = st.number_input("AUM Finance (EUR)",
                            value=float(row_data.get("funded_aum",0.0)),
                            min_value=0.0, step=1_000_000.0, format="%.0f")
                    r3c1, r3c2, r3c3, r3c4 = st.columns(4)
                    with r3c1:
                        ropts = [""] + RAISONS_PERTE
                        cur_r = str(row_data.get("raison_perte") or "")
                        ri    = ropts.index(cur_r) if cur_r in ropts else 0
                        lbl_r = "Raison (obligatoire)" if new_statut in ("Lost","Paused") else "Raison"
                        new_raison = st.selectbox(lbl_r, ropts, index=ri)
                    with r3c2:
                        new_conc = st.text_input("Concurrent",
                            value=str(row_data.get("concurrent_choisi") or ""))
                    with r3c3:
                        nad = row_data.get("next_action_date")
                        if not isinstance(nad, date): nad = date.today() + timedelta(days=14)
                        new_nad = st.date_input("Prochaine Action", value=nad)
                    with r3c4:
                        new_sales = st.text_input("Commercial",
                            value=str(row_data.get("sales_owner") or "Non assigne"))
                    sub = st.form_submit_button("Sauvegarder", use_container_width=True)
                if sub:
                    ok, msg = db.update_pipeline_row({
                        "id": pipeline_id, "fonds": new_fonds, "statut": new_statut,
                        "target_aum_initial": new_target, "revised_aum": new_revised,
                        "funded_aum": new_funded, "raison_perte": new_raison,
                        "concurrent_choisi": new_conc, "next_action_date": new_nad,
                        "sales_owner": new_sales})
                    if ok:
                        st.success("Deal mis a jour — {} / {}".format(new_statut, fmt_m(new_funded)))
                        st.rerun()
                    else:
                        st.error(msg)
            df_audit = db.get_audit_log(pipeline_id)
            if not df_audit.empty:
                with st.expander("Historique modifications — Deal #{}".format(pipeline_id)):
                    st.dataframe(df_audit, use_container_width=True, hide_index=True,
                                 height=min(220, 46 + len(df_audit) * 36))
    else:
        st.markdown(
            '<div style="background:#001c4b04;border:1px dashed #001c4b20;'
            'padding:20px;text-align:center;margin-top:10px;">'
            '<div style="color:{};font-weight:600;font-size:0.85rem;">Selectionnez un deal dans le tableau</div>'
            '<div style="color:#888;font-size:0.75rem;margin-top:2px;">'
            "Le formulaire d'edition s'affichera ici</div>"
            '</div>'.format(MARINE), unsafe_allow_html=True)

    st.divider()
    st.markdown("#### Comparaison AUM par Deal Actif")
    df_viz = db.get_pipeline_with_clients()
    df_viz = df_viz[
        (df_viz["target_aum_initial"] > 0) &
        (df_viz["statut"].isin(["Funded","Soft Commit","Due Diligence","Initial Pitch"]))
    ].head(10)
    if not df_viz.empty:
        fig_viz = go.Figure()
        for label, col, color in [
            ("AUM Cible",  "target_aum_initial", GRIS),
            ("AUM Revise", "revised_aum",         B_MID),
            ("AUM Finance","funded_aum",           MARINE)]:
            fig_viz.add_trace(go.Bar(name=label, x=df_viz["nom_client"].str[:18].tolist(),
                                     y=df_viz[col].tolist(), marker_color=color,
                                     marker_line_color=BLANC, marker_line_width=0.5))
        fig_viz.update_layout(
            height=320, paper_bgcolor=BLANC, plot_bgcolor=BLANC,
            font_color=MARINE, barmode="group", bargap=0.22, bargroupgap=0.06,
            legend_bgcolor=BLANC, legend_font_size=10,
            xaxis_tickangle=-20, xaxis_showgrid=False,
            yaxis_showgrid=True, yaxis_gridcolor=GRIS,
            margin=dict(l=10, r=10, t=20, b=10))
        st.plotly_chart(fig_viz, use_container_width=True, config={"displayModeBar": False})

    df_lp = db.get_pipeline_with_clients()
    df_lp = df_lp[df_lp["statut"].isin(["Lost","Paused"])].copy()
    if not df_lp.empty:
        st.divider()
        st.markdown("#### Deals Perdus / En Pause")
        df_lp2 = df_lp[["nom_client","fonds","statut","target_aum_initial",
                         "raison_perte","concurrent_choisi","sales_owner"]].copy()
        df_lp2["target_aum_initial"] = df_lp2["target_aum_initial"].apply(fmt_m)
        st.dataframe(df_lp2, use_container_width=True, hide_index=True)


# ============================================================================
# ONGLET 3 — EXECUTIVE DASHBOARD
# Zéro st.button visible — tout est un lien HTML pur via clickable()
# ============================================================================
with tab_dash:
    st.markdown('<div class="section-title">Executive Dashboard</div>',
                unsafe_allow_html=True)

    kpis = db.get_kpis()
    nb_lost_paused = kpis["nb_lost"] + kpis.get("nb_paused", 0)

    # ---- KPI Cards — liens purs, zéro bouton ----
    kc1, kc2, kc3, kc4 = st.columns(4, gap="small")

    with kc1:
        clickable(
            '<div class="kpi-card">'
            '<div class="kpi-label">AUM Finance Total</div>'
            '<div class="kpi-value">{}</div>'
            '<div class="kpi-sub" style="color:#019ee1;">{} deal(s) Funded</div>'
            '</div>'.format(fmt_m(kpis["total_funded"]), kpis["nb_funded"]),
            "funded")

    with kc2:
        clickable(
            '<div class="kpi-card">'
            '<div class="kpi-label">Pipeline Actif</div>'
            '<div class="kpi-value">{}</div>'
            '<div class="kpi-sub" style="color:#019ee1;">{} deals en cours</div>'
            '</div>'.format(fmt_m(kpis["pipeline_actif"]), kpis["nb_deals_actifs"]),
            "pipeline")

    with kc3:
        # Pas de drill-down — pas de lien
        st.markdown(
            '<div class="kpi-card kpi-card-static">'
            '<div class="kpi-label">Taux Conversion</div>'
            '<div class="kpi-value">{:.1f}%</div>'
            '<div class="kpi-sub">{} funded / {} lost</div>'
            '</div>'.format(kpis["taux_conversion"], kpis["nb_funded"], kpis["nb_lost"]),
            unsafe_allow_html=True)

    with kc4:
        if nb_lost_paused > 0:
            clickable(
                '<div class="kpi-card">'
                '<div class="kpi-label">Lost / Paused</div>'
                '<div class="kpi-value">{}</div>'
                '<div class="kpi-sub" style="color:#019ee1;">{} deals</div>'
                '</div>'.format(nb_lost_paused, nb_lost_paused),
                "lost")
        else:
            st.markdown(
                '<div class="kpi-card kpi-card-static">'
                '<div class="kpi-label">Lost / Paused</div>'
                '<div class="kpi-value">0</div>'
                '<div class="kpi-sub">Aucun deal perdu</div>'
                '</div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ---- Pastilles statut — liens purs ----
    statut_order = [s for s in STATUTS if kpis["statut_repartition"].get(s, 0) > 0]
    if statut_order:
        bcols = st.columns(len(statut_order), gap="small")
        for col, s in zip(bcols, statut_order):
            c_hex = STATUT_COLORS.get(s, GRIS)
            count = kpis["statut_repartition"][s]
            with col:
                clickable(
                    '<div class="statut-pill" style="background:{c}16;border:1px solid {c}44;">'
                    '<div style="font-size:0.61rem;color:{marine};font-weight:700;'
                    'text-transform:uppercase;">{s}</div>'
                    '<div style="font-size:1.3rem;font-weight:800;color:{c};">{n}</div>'
                    '</div>'.format(c=c_hex, marine=MARINE, s=s, n=count),
                    "statut_{}".format(s))

    st.divider()

    # ---- Alertes retard — liens purs ----
    df_overdue = db.get_overdue_actions()
    if not df_overdue.empty:
        st.markdown(
            "<div style='font-size:0.84rem;font-weight:700;color:#001c4b;"
            "border-left:3px solid #f07d00;padding:4px 10px;margin-bottom:4px;'>"
            "{} action(s) en retard</div>".format(len(df_overdue)),
            unsafe_allow_html=True)
        today = date.today()
        for idx, row in df_overdue.iterrows():
            nad = row.get("next_action_date")
            days_late = (today - nad).days if isinstance(nad, date) else 0
            nad_str   = nad.isoformat() if isinstance(nad, date) else "—"
            owner     = str(row.get("sales_owner","")) or ""
            clickable(
                '<div class="alert-overdue">'
                '<b>{client}</b> — {fonds}'
                ' <span style="color:{ciel};font-weight:600;">({statut})</span>'
                ' — Prevue le <b>{nad}</b>'
                ' &nbsp;<span class="badge-retard">RETARD +{days}j</span>'
                '{owner_part}'
                '</div>'.format(
                    client=row["nom_client"], fonds=row["fonds"], ciel=CIEL,
                    statut=row["statut"], nad=nad_str, days=days_late,
                    owner_part=" — <b>{}</b>".format(owner) if owner else ""),
                "overdue")

    # ---- Graphiques ----
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
        st.markdown("#### Par Region")
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
                                    yaxis_autorange="reversed",
                                    margin=dict(l=10, r=60, t=36, b=10))
            st.plotly_chart(fig_fonds, use_container_width=True, config={"displayModeBar": False})

    st.divider()
    st.markdown("#### Top Deals — AUM Finance")
    df_funded_top = (db.get_pipeline_with_clients()
                     .query("statut == 'Funded'")
                     .sort_values("funded_aum", ascending=False)
                     .head(10))
    if not df_funded_top.empty:
        max_f = float(df_funded_top["funded_aum"].max())
        for i, (_, row) in enumerate(df_funded_top.iterrows()):
            val  = float(row["funded_aum"])
            pct  = val / max_f * 100 if max_f > 0 else 0
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
                '<div style="font-size:0.61rem;color:#999;">finance</div>'
                '</div></div>'.format(
                    rank=i+1, ciel=CIEL, marine=MARINE, gris=GRIS,
                    client=row["nom_client"], fonds=row["fonds"],
                    type=row["type_client"], region=row["region"],
                    barc=bar_c, pct=pct, aum=fmt_m(val)), unsafe_allow_html=True)


# ============================================================================
# ONGLET 4 — SALES TRACKING
# ============================================================================
with tab_sales:
    st.markdown('<div class="section-title">Sales Tracking — Suivi par Commercial</div>',
                unsafe_allow_html=True)
    df_sm = db.get_sales_metrics()
    df_na = db.get_next_actions_by_sales(days_ahead=30)

    if df_sm.empty:
        st.info("Aucune donnee de pipeline disponible.")
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
        st.markdown("#### AUM Finance vs Pipeline par Commercial")
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
        st.markdown("#### Prochaines Actions — 30 jours")
        if df_na.empty:
            st.info("Aucune action planifiee dans les 30 prochains jours.")
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
                        statut=row.get("statut",""), aum=fmt_m(float(row.get("revised_aum",0) or 0)),
                        owner=row.get("sales_owner","")), unsafe_allow_html=True)


# ============================================================================
# ONGLET 5 — PERFORMANCE & NAV
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
        st.markdown("#### Demonstration")
        if st.button("Generer un fichier NAV de demonstration", use_container_width=True):
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
                st.error("Aucune donnee valide."); st.stop()

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
                fonds_sel_nav = st.multiselect("Fonds a afficher", fonds_list,
                                               default=fonds_list[:min(5, len(fonds_list))])
            with ff2: d_debut = st.date_input("Depuis", value=d_min.date())
            with ff3: d_fin   = st.date_input("Jusqu'au", value=d_max.date())

            if not fonds_sel_nav:
                st.warning("Selectionnez au moins un fonds."); st.stop()

            mask = (df_nav["Fonds"].isin(fonds_sel_nav) &
                    (df_nav["Date"].dt.date >= d_debut) &
                    (df_nav["Date"].dt.date <= d_fin))
            df_fn = df_nav[mask].copy()
            if df_fn.empty:
                st.warning("Aucune donnee pour la periode."); st.stop()

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
                    "Perf 1M (%)":    round(p1m,  2) if not np.isnan(p1m)  else None,
                    "Perf YTD (%)":   round(pytd, 2) if not np.isnan(pytd) else None,
                    "Perf Periode (%)": round(pp, 2) if not np.isnan(pp)   else None})

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
                    '<th style="padding:8px 12px;text-align:right;">Perf Periode</th>'
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
            with st.expander("Details"):
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
