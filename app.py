# =============================================================================
# app.py — CRM & Reporting — Amundi Research Grade Edition
# Charte : Marine #002D54 | Orange #FF4F00 | Fond Blanc #FFFFFF
# Standards : Zero emoji, Plotly interactif, modales st.dialog, M EUR
# Architecture :
#   - Plotly UNIQUEMENT pour les graphiques Streamlit (interactivite)
#   - Matplotlib reservé au pdf_generator.py (PNG statique)
#   - KPIs cliquables -> st.dialog (drill-down)
#   - Sales Tracking : charge par commercial + Next Actions priorisees
#   - Audit trail sobre sous chaque deal Pipeline
# Lancement : streamlit run app.py
# =============================================================================

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from datetime import date, timedelta
import io
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import database as db
import pdf_generator as pdf_gen

# ---------------------------------------------------------------------------
# PALETTE & CONSTANTES CHARTE AMUNDI
# ---------------------------------------------------------------------------

C_MARINE     = "#002D54"
C_ORANGE     = "#FF4F00"   # UNIQUEMENT : boutons save, retards, alertes critiques
C_BLANC      = "#FFFFFF"
C_GRIS       = "#E8E8E8"
C_GRIS_TXT   = "#555555"
C_BLEU_MID   = "#1A5E8A"
C_BLEU_PAL   = "#4A8FBD"
C_BLEU_DEP   = "#003F7A"

# Palette graphiques — bleu marine et variations institutionnelles
PLOTLY_PALETTE = [
    C_BLEU_MID, C_MARINE, C_BLEU_PAL, C_BLEU_DEP,
    "#2C7FB8", "#004F8C", "#6BAED6", "#08519C",
    "#9ECAE1", "#003060",
]

STATUT_COLORS = {
    "Funded":        "#1A5E8A",
    "Soft Commit":   "#2C7FB8",
    "Due Diligence": "#004F8C",
    "Initial Pitch": "#4A8FBD",
    "Prospect":      "#9ECAE1",
    "Lost":          "#AAAAAA",
    "Paused":        "#C0C0C0",
    "Redeemed":      "#B8B8D0",
}

TYPES_CLIENT      = ["IFA", "Wholesale", "Instit", "Family Office"]
REGIONS           = db.REGIONS_REFERENTIEL
FONDS             = db.FONDS_REFERENTIEL
STATUTS           = ["Prospect", "Initial Pitch", "Due Diligence",
                     "Soft Commit", "Funded", "Paused", "Lost", "Redeemed"]
STATUTS_ACTIFS    = ["Prospect", "Initial Pitch", "Due Diligence", "Soft Commit"]
RAISONS_PERTE     = ["Pricing", "Track Record", "Macro", "Competitor", "Autre"]
TYPES_INTERACTION = ["Call", "Meeting", "Email", "Roadshow", "Conference", "Autre"]


# ---------------------------------------------------------------------------
# CONFIG STREAMLIT
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="CRM — Asset Management",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ---------------------------------------------------------------------------
# CSS — CHARTE AMUNDI STRICTE
# Orange uniquement : boutons Sauvegarder, badges retard, alertes critiques
# ---------------------------------------------------------------------------

st.markdown(f"""
<style>
  /* Base */
  .stApp, .main .block-container {{
      background-color:{C_BLANC};
      color:{C_MARINE};
      font-family:'Segoe UI', Arial, sans-serif;
  }}

  /* Sidebar */
  [data-testid="stSidebar"] {{ background-color:{C_MARINE}; }}
  [data-testid="stSidebar"] * {{ color:{C_BLANC} !important; }}
  [data-testid="stSidebar"] h1,
  [data-testid="stSidebar"] h2,
  [data-testid="stSidebar"] h3 {{
      color:{C_BLEU_PAL} !important;
  }}

  /* En-tete application */
  .crm-header {{
      background:{C_MARINE};
      padding:16px 24px;
      border-radius:6px;
      margin-bottom:16px;
      border-bottom:3px solid {C_ORANGE};
  }}
  .crm-header h1 {{
      color:{C_BLANC} !important;
      margin:0;
      font-size:1.45rem;
      font-weight:700;
      letter-spacing:-0.2px;
  }}
  .crm-header p {{
      color:{C_BLEU_PAL};
      margin:3px 0 0 0;
      font-size:0.80rem;
  }}

  /* KPI Cards — fond marine, texte blanc */
  .kpi-card {{
      background:{C_MARINE};
      padding:14px 10px;
      border-radius:6px;
      text-align:center;
      border:1px solid {C_MARINE}55;
      cursor:pointer;
      transition:background 0.18s;
  }}
  .kpi-card:hover {{
      background:{C_BLEU_MID};
  }}
  .kpi-label {{
      font-size:0.66rem;
      color:{C_BLEU_PAL};
      text-transform:uppercase;
      letter-spacing:0.8px;
      margin-bottom:5px;
      font-weight:600;
  }}
  .kpi-value {{ font-size:1.45rem; font-weight:800; color:{C_BLANC}; }}
  .kpi-sub   {{ font-size:0.63rem; color:{C_GRIS}; margin-top:3px; }}

  /* Titres de sections */
  .section-title {{
      font-size:0.92rem;
      font-weight:700;
      color:{C_MARINE};
      border-bottom:2px solid {C_MARINE}22;
      padding-bottom:4px;
      margin:14px 0 9px 0;
  }}

  /* Badge retard — Orange Amundi */
  .badge-retard {{
      display:inline-block;
      background:{C_ORANGE};
      color:{C_BLANC};
      border-radius:3px;
      padding:1px 7px;
      font-size:0.68rem;
      font-weight:700;
      letter-spacing:0.4px;
  }}

  /* Badge statut */
  .perimetre-badge {{
      display:inline-block;
      background:{C_BLEU_PAL}18;
      border:1px solid {C_BLEU_PAL}44;
      border-radius:4px;
      padding:2px 8px;
      font-size:0.72rem;
      color:{C_MARINE};
      font-weight:600;
      margin:2px;
  }}

  /* Panneau de detail */
  .detail-panel {{
      background:#F4F8FC;
      border:1px solid {C_MARINE}18;
      border-radius:8px;
      padding:16px 18px 12px 18px;
      margin-top:12px;
  }}

  /* Alerte action en retard */
  .alert-overdue {{
      background:#FFF5F0;
      border-left:3px solid {C_ORANGE};
      border-radius:0 4px 4px 0;
      padding:6px 11px;
      margin:3px 0;
      font-size:0.78rem;
      color:{C_MARINE};
  }}

  /* Card commercial */
  .sales-card {{
      background:{C_BLANC};
      border:1px solid {C_MARINE}12;
      border-radius:6px;
      padding:13px;
      border-top:3px solid {C_MARINE};
  }}
  .sales-card-name {{
      font-size:0.88rem;
      font-weight:700;
      color:{C_MARINE};
      margin-bottom:8px;
      padding-bottom:5px;
      border-bottom:1px solid {C_GRIS};
  }}
  .sales-metric     {{ font-size:0.70rem; color:#666; margin-bottom:2px; }}
  .sales-metric-val {{ font-size:0.95rem; font-weight:700; color:{C_MARINE}; }}
  .sales-metric-acc {{ font-size:0.95rem; font-weight:700; color:{C_BLEU_MID}; }}

  /* Onglets */
  .stTabs [data-baseweb="tab-list"] {{
      background:{C_BLANC};
      border-bottom:2px solid {C_MARINE}14;
      gap:2px;
  }}
  .stTabs [data-baseweb="tab"] {{
      color:{C_MARINE};
      font-weight:600;
      font-size:0.81rem;
      padding:7px 15px;
      border-radius:4px 4px 0 0;
      background:{C_MARINE}06;
  }}
  .stTabs [aria-selected="true"] {{
      background:{C_MARINE} !important;
      color:{C_BLANC} !important;
  }}

  /* Boutons — Marine standard */
  .stButton > button {{
      background:{C_MARINE};
      color:{C_BLANC};
      border:none;
      border-radius:4px;
      font-weight:600;
      padding:6px 15px;
      font-size:0.81rem;
      transition:background 0.15s;
  }}
  .stButton > button:hover {{
      background:{C_BLEU_MID};
      color:{C_BLANC};
  }}

  /* Bouton Sauvegarder — Orange Amundi (element critique) */
  [data-testid="stFormSubmitButton"] > button {{
      background:{C_ORANGE} !important;
      color:{C_BLANC} !important;
      border:none;
      font-weight:700;
  }}
  [data-testid="stFormSubmitButton"] > button:hover {{
      background:#D94200 !important;
  }}

  /* Labels formulaires */
  .stSelectbox label, .stTextInput label, .stNumberInput label,
  .stDateInput label, .stTextArea label, .stRadio label {{
      color:{C_MARINE} !important;
      font-weight:600;
      font-size:0.78rem;
  }}

  /* Titres */
  h1, h2, h3, h4 {{ color:{C_MARINE} !important; }}
  hr {{ border-color:{C_MARINE}10; }}
  code {{
      background:{C_MARINE}08;
      color:{C_MARINE};
      border-radius:3px;
      font-size:0.82em;
  }}

  /* Tableau de bord commercial pipeline */
  .pipeline-hint {{
      background:{C_BLEU_PAL}0D;
      border-left:2px solid {C_MARINE};
      border-radius:0 4px 4px 0;
      padding:6px 11px;
      font-size:0.78rem;
      color:{C_MARINE};
      margin-bottom:8px;
  }}

  /* Audit trail */
  .audit-row {{
      padding:5px 10px;
      border-bottom:1px solid {C_GRIS};
      font-size:0.76rem;
      color:{C_MARINE};
  }}
  .audit-row:last-child {{ border-bottom:none; }}

  /* Separateur de section */
  .section-sep {{
      border:none;
      border-top:1px solid {C_MARINE}14;
      margin:14px 0;
  }}
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
# HELPERS
# ---------------------------------------------------------------------------

def fmt_m(v) -> str:
    """
    Formatage institutionnel Amundi : Millions EUR (1 decimale).
    Force l'affichage en M EUR ou Md EUR. Zero symbole $.
    """
    try:
        fv = float(v)
    except (TypeError, ValueError):
        return "—"
    if fv >= 1_000_000_000:
        return f"{fv / 1_000_000_000:.1f} Md EUR"
    if fv >= 1_000_000:
        return f"{fv / 1_000_000:.1f} M EUR"
    if fv >= 1_000:
        return f"{fv / 1_000:.0f} k EUR"
    return f"{fv:.0f} EUR"


def kpi_card(label: str, value: str, sub: str = "") -> str:
    return (
        f'<div class="kpi-card">'
        f'<div class="kpi-label">{label}</div>'
        f'<div class="kpi-value">{value}</div>'
        + (f'<div class="kpi-sub">{sub}</div>' if sub else "")
        + "</div>"
    )


def statut_badge(statut: str) -> str:
    colors = {
        "Funded":        (C_BLEU_MID,   C_BLANC),
        "Soft Commit":   ("#2C7FB8",    C_BLANC),
        "Due Diligence": ("#004F8C",    C_BLANC),
        "Initial Pitch": ("#4A8FBD",    C_BLANC),
        "Prospect":      (f"{C_MARINE}18", C_MARINE),
        "Lost":          (C_GRIS,       "#555"),
        "Paused":        ("#D0D0D0",    C_MARINE),
        "Redeemed":      ("#C8D8E8",    C_MARINE),
    }
    bg, fg = colors.get(statut, (C_GRIS, "#555"))
    return (
        f'<span style="padding:2px 9px;border-radius:9px;'
        f'font-size:0.72rem;font-weight:600;'
        f'background:{bg};color:{fg};">{statut}</span>'
    )


def _plotly_layout(title: str = "", height: int = 360) -> dict:
    """Parametres layout Plotly communs — charte Amundi."""
    return dict(
        title=dict(
            text=title,
            font=dict(size=13, color=C_MARINE, family="Segoe UI, Arial"),
            x=0, xanchor="left"
        ),
        height=height,
        paper_bgcolor=C_BLANC,
        plot_bgcolor=C_BLANC,
        margin=dict(l=10, r=10, t=40 if title else 10, b=10),
        font=dict(family="Segoe UI, Arial", size=11, color=C_MARINE),
        legend=dict(
            bgcolor=C_BLANC,
            bordercolor=C_GRIS,
            borderwidth=1,
            font=dict(size=10, color=C_MARINE),
        ),
        hoverlabel=dict(
            bgcolor=C_MARINE,
            font=dict(color=C_BLANC, size=11),
            bordercolor=C_MARINE,
        ),
    )


# ---------------------------------------------------------------------------
# MODALES DRILL-DOWN (st.dialog)
# ---------------------------------------------------------------------------

@st.dialog("Detail — Deals Finances")
def modal_funded(fonds_filter=None):
    df = db.get_funded_deals_detail(fonds_filter)
    if df.empty:
        st.info("Aucun deal Funded dans le perimetre.")
        return
    df_disp = df.copy()
    df_disp["AUM Finance"] = df_disp["AUM Finance"].apply(fmt_m)
    st.dataframe(
        df_disp,
        use_container_width=True,
        hide_index=True,
        height=min(420, 46 + len(df_disp) * 36),
        column_config={
            "Client":       st.column_config.TextColumn("Client"),
            "Fonds":        st.column_config.TextColumn("Fonds"),
            "Type":         st.column_config.TextColumn("Type", width="small"),
            "Region":       st.column_config.TextColumn("Region", width="small"),
            "AUM Finance":  st.column_config.TextColumn("AUM Finance"),
            "Commercial":   st.column_config.TextColumn("Commercial"),
        }
    )


@st.dialog("Detail — Pipeline Actif")
def modal_pipeline_actif(fonds_filter=None):
    df = db.get_active_deals_detail(fonds_filter)
    if df.empty:
        st.info("Aucun deal actif dans le perimetre.")
        return
    df_disp = df.copy()
    df_disp["AUM Revise"] = df_disp["AUM Revise"].apply(fmt_m)
    if "Prochaine Action" in df_disp.columns:
        df_disp["Prochaine Action"] = df_disp["Prochaine Action"].apply(
            lambda v: v.isoformat() if isinstance(v, date) else str(v or "—")
        )
    st.dataframe(
        df_disp,
        use_container_width=True,
        hide_index=True,
        height=min(420, 46 + len(df_disp) * 36),
    )


@st.dialog("Detail — Deals Perdus et En Pause")
def modal_lost(fonds_filter=None):
    df = db.get_lost_deals_detail(fonds_filter)
    if df.empty:
        st.info("Aucun deal Lost/Paused dans le perimetre.")
        return
    df_disp = df.copy()
    df_disp["AUM Cible"] = df_disp["AUM Cible"].apply(fmt_m)
    st.dataframe(
        df_disp,
        use_container_width=True,
        hide_index=True,
        height=min(420, 46 + len(df_disp) * 36),
        column_config={
            "Client":     st.column_config.TextColumn("Client"),
            "Fonds":      st.column_config.TextColumn("Fonds"),
            "Statut":     st.column_config.TextColumn("Statut", width="small"),
            "AUM Cible":  st.column_config.TextColumn("AUM Cible"),
            "Raison":     st.column_config.TextColumn("Raison"),
            "Concurrent": st.column_config.TextColumn("Concurrent"),
            "Commercial": st.column_config.TextColumn("Commercial"),
        }
    )


# ---------------------------------------------------------------------------
# SIDEBAR — PERIMETRE D'EXPORT
# ---------------------------------------------------------------------------

with st.sidebar:
    st.markdown(f"""
    <div style="text-align:center;padding:6px 0 14px 0;">
        <div style="font-size:0.92rem;font-weight:800;color:{C_BLANC};
                    letter-spacing:0.3px;">
            CRM Asset Management
        </div>
        <div style="font-size:0.66rem;color:{C_GRIS};margin-top:3px;">
            {date.today().strftime("%d %B %Y")}
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.divider()

    # Apercu global
    kpis_global = db.get_kpis()
    st.markdown(
        f'<div style="color:{C_BLEU_PAL};font-size:0.68rem;font-weight:700;'
        f'text-transform:uppercase;letter-spacing:1px;margin-bottom:5px;">'
        f'Apercu Global</div>',
        unsafe_allow_html=True
    )
    for lbl, val in [
        ("AUM Finance Total", fmt_m(kpis_global["total_funded"])),
        ("Pipeline Actif",    fmt_m(kpis_global["pipeline_actif"])),
    ]:
        st.markdown(f"""
        <div style="background:{C_BLANC}14;padding:8px;border-radius:5px;
                    margin-bottom:5px;">
            <div style="font-size:0.64rem;color:{C_GRIS};">{lbl}</div>
            <div style="font-size:1.10rem;font-weight:800;color:{C_BLANC};">
                {val}</div>
        </div>""", unsafe_allow_html=True)

    st.divider()

    st.markdown(
        f'<div style="color:{C_BLEU_PAL};font-size:0.68rem;font-weight:700;'
        f'text-transform:uppercase;letter-spacing:1px;margin-bottom:5px;">'
        f'Export PDF</div>',
        unsafe_allow_html=True
    )

    fonds_perimetre = st.multiselect(
        "Perimetre de l'export",
        options=FONDS,
        default=FONDS,
        help=(
            "Selectionnez les fonds a inclure dans le rapport PDF. "
            "KPIs, graphiques et tableaux seront recalcules sur ce perimetre."
        ),
        key="fonds_perimetre_select"
    )

    _filtre_effectif = (
        fonds_perimetre
        if (fonds_perimetre and len(fonds_perimetre) < len(FONDS))
        else None
    )

    mode_comex = st.toggle(
        "Mode Comex — Anonymisation",
        value=False,
        help="Remplace les noms clients par Type — Region dans le PDF."
    )

    if fonds_perimetre and len(fonds_perimetre) < len(FONDS):
        badges = " ".join(
            f'<span class="perimetre-badge">{f}</span>'
            for f in fonds_perimetre
        )
        st.markdown(f"<div>{badges}</div>", unsafe_allow_html=True)
    elif not fonds_perimetre:
        st.warning("Selectionnez au moins un fonds.")

    perf_data_pdf   = st.session_state.get("perf_data", None)
    nav_base100_pdf = st.session_state.get("nav_base100", None)
    if perf_data_pdf is not None:
        nb_fonds_nav = len(perf_data_pdf)
        st.caption(f"Donnees NAV : {nb_fonds_nav} fonds charges.")

    if not fonds_perimetre:
        st.button("Generer le rapport PDF",
                  disabled=True, use_container_width=True)
    elif st.button("Generer le rapport PDF", use_container_width=True):
        with st.spinner("Generation du rapport en cours..."):
            try:
                pipeline_pdf   = db.get_pipeline_with_clients(
                                     fonds_filter=_filtre_effectif)
                kpis_pdf       = db.get_kpis(fonds_filter=_filtre_effectif)
                aum_region_pdf = db.get_aum_by_region(
                                     fonds_filter=_filtre_effectif)

                pf_pdf = perf_data_pdf
                nb_pdf = nav_base100_pdf
                if pf_pdf is not None and _filtre_effectif:
                    pf_pdf = pf_pdf[pf_pdf["Fonds"].isin(_filtre_effectif)] \
                             if "Fonds" in pf_pdf.columns else pf_pdf
                if nb_pdf is not None and _filtre_effectif:
                    cols_k = [c for c in nb_pdf.columns
                              if c in _filtre_effectif]
                    nb_pdf = nb_pdf[cols_k] if cols_k else None

                pdf_bytes = pdf_gen.generate_pdf(
                    pipeline_df    = pipeline_pdf,
                    kpis           = kpis_pdf,
                    aum_by_region  = aum_region_pdf,
                    mode_comex     = mode_comex,
                    perf_data      = pf_pdf,
                    nav_base100_df = nb_pdf,
                    fonds_perimetre= fonds_perimetre,
                )
                fname = (
                    f"report_comex_{date.today().isoformat()}.pdf"
                    if mode_comex
                    else f"report_{date.today().isoformat()}.pdf"
                )
                st.download_button(
                    "Telecharger le rapport",
                    data=pdf_bytes,
                    file_name=fname,
                    mime="application/pdf",
                    use_container_width=True,
                )
                st.success("Rapport genere.")
            except Exception as e:
                st.error(f"Erreur : {e}")

    st.divider()
    st.caption("Version 5.0 — Amundi Research Grade Edition")


# ---------------------------------------------------------------------------
# EN-TETE
# ---------------------------------------------------------------------------

st.markdown(f"""
<div class="crm-header">
    <h1>CRM &amp; Reporting — Asset Management</h1>
    <p>Pipeline commercial &middot; Suivi des mandats &middot;
       Reporting executif &middot; Performance NAV</p>
</div>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# ONGLETS
# ---------------------------------------------------------------------------

tab_ingest, tab_pipeline, tab_dash, tab_sales, tab_perf = st.tabs([
    "Data Ingestion",
    "Pipeline Management",
    "Executive Dashboard",
    "Sales Tracking",
    "Performance et NAV",
])


# ============================================================================
# ONGLET 1 : DATA INGESTION
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
                sub_c = st.form_submit_button(
                    "Enregistrer le client", use_container_width=True)
            if sub_c:
                if not nom_client.strip():
                    st.error("Nom du client obligatoire.")
                else:
                    try:
                        db.add_client(nom_client.strip(), type_client, region)
                        st.success(f"Client {nom_client} ajoute.")
                    except Exception as e:
                        st.warning(
                            "Ce client existe deja."
                            if "UNIQUE" in str(e) else f"Erreur : {e}"
                        )

        with st.expander("Ajouter un Deal Pipeline", expanded=True):
            clients_dict = db.get_client_options()
            sales_owners = db.get_sales_owners() or ["Non assigne"]
            if not clients_dict:
                st.info("Ajoutez d'abord un client.")
            else:
                with st.form("form_deal", clear_on_submit=True):
                    ca, cb = st.columns(2)
                    with ca:
                        client_sel  = st.selectbox(
                            "Client", list(clients_dict.keys()))
                        fonds_sel   = st.selectbox("Fonds", FONDS)
                        statut_sel  = st.selectbox("Statut", STATUTS)
                        owner_opts  = sales_owners + ["Nouveau commercial..."]
                        owner_sel   = st.selectbox("Commercial", owner_opts)
                        if owner_sel == "Nouveau commercial...":
                            owner_input = st.text_input(
                                "Nom du nouveau commercial")
                        else:
                            owner_input = owner_sel
                    with cb:
                        target_aum  = st.number_input(
                            "AUM Cible (EUR)",
                            min_value=0, step=1_000_000)
                        revised_aum = st.number_input(
                            "AUM Revise (EUR)",
                            min_value=0, step=1_000_000)
                        funded_aum  = st.number_input(
                            "AUM Finance (EUR)",
                            min_value=0, step=1_000_000)

                    raison_perte, concurrent = "", ""
                    if statut_sel in ("Lost", "Paused"):
                        cc, cd = st.columns(2)
                        with cc:
                            raison_perte = st.selectbox(
                                "Raison", RAISONS_PERTE)
                        with cd:
                            concurrent = st.text_input("Concurrent")

                    next_action = st.date_input(
                        "Prochaine Action",
                        value=date.today() + timedelta(days=14))
                    sub_d = st.form_submit_button(
                        "Enregistrer le deal", use_container_width=True)

                if sub_d:
                    if statut_sel in ("Lost", "Paused") and not raison_perte:
                        st.error("Raison obligatoire pour ce statut.")
                    else:
                        db.add_pipeline_entry(
                            clients_dict[client_sel], fonds_sel, statut_sel,
                            float(target_aum), float(revised_aum),
                            float(funded_aum),
                            raison_perte, concurrent,
                            next_action.isoformat(),
                            owner_input.strip() or "Non assigne"
                        )
                        st.success(
                            f"Deal {fonds_sel} / {client_sel} enregistre.")

        with st.expander("Enregistrer une Activite"):
            clients_dict2 = db.get_client_options()
            if clients_dict2:
                with st.form("form_act", clear_on_submit=True):
                    ce, cf = st.columns(2)
                    with ce:
                        act_client = st.selectbox(
                            "Client", list(clients_dict2.keys()))
                        act_type   = st.selectbox(
                            "Type d'interaction", TYPES_INTERACTION)
                    with cf:
                        act_date  = st.date_input("Date", value=date.today())
                        act_notes = st.text_area("Notes", height=68)
                    sub_a = st.form_submit_button(
                        "Enregistrer", use_container_width=True)
                if sub_a:
                    db.add_activity(
                        clients_dict2[act_client],
                        act_date.isoformat(), act_notes, act_type)
                    st.success(f"Activite enregistree pour {act_client}.")

    with col_import:
        st.markdown("#### Import CSV / Excel — Upsert")
        import_type = st.radio(
            "Table cible", ["Clients", "Pipeline"], horizontal=True)
        if import_type == "Clients":
            st.info("Colonnes attendues : nom_client, type_client, region")
        else:
            st.info(
                "Colonnes attendues : nom_client, fonds, statut, "
                "target_aum_initial, revised_aum, funded_aum, "
                "raison_perte, concurrent_choisi, next_action_date, "
                "sales_owner"
            )

        uploaded_file = st.file_uploader(
            "Fichier CSV ou Excel", type=["csv", "xlsx", "xls"])
        if uploaded_file:
            try:
                df_imp = (pd.read_csv(uploaded_file)
                          if uploaded_file.name.endswith(".csv")
                          else pd.read_excel(uploaded_file))
                st.dataframe(
                    df_imp.head(5),
                    use_container_width=True, height=145)
                st.caption(f"{len(df_imp)} ligne(s) detectees")
                if st.button("Lancer l'import", use_container_width=True):
                    fn = (db.upsert_clients_from_df
                          if import_type == "Clients"
                          else db.upsert_pipeline_from_df)
                    ins, upd = fn(df_imp)
                    st.success(
                        f"Import termine : {ins} cree(s), {upd} mis a jour.")
            except Exception as e:
                st.error(f"Erreur de lecture : {e}")

        st.divider()
        st.markdown("#### Dernieres Activites")
        df_act = db.get_activities()
        if not df_act.empty:
            st.dataframe(
                df_act[["nom_client", "date",
                         "type_interaction", "notes"]].head(10),
                use_container_width=True,
                height=250,
                hide_index=True,
                column_config={
                    "nom_client":       st.column_config.TextColumn("Client"),
                    "date":             st.column_config.TextColumn("Date"),
                    "type_interaction": st.column_config.TextColumn("Type"),
                    "notes":            st.column_config.TextColumn("Notes"),
                })


# ============================================================================
# ONGLET 2 : PIPELINE MANAGEMENT — MASTER-DETAIL + AUDIT TRAIL
# ============================================================================
with tab_pipeline:
    st.markdown('<div class="section-title">Pipeline Management</div>',
                unsafe_allow_html=True)

    with st.expander("Filtres", expanded=False):
        fc1, fc2, fc3 = st.columns(3)
        with fc1:
            filt_statuts = st.multiselect(
                "Statuts", STATUTS, default=STATUTS_ACTIFS)
        with fc2:
            filt_fonds   = st.multiselect("Fonds", FONDS)
        with fc3:
            filt_regions = st.multiselect("Regions", REGIONS)

    df_pipe = db.get_pipeline_with_clients()
    df_view = df_pipe.copy()
    if filt_statuts:  df_view = df_view[df_view["statut"].isin(filt_statuts)]
    if filt_fonds:    df_view = df_view[df_view["fonds"].isin(filt_fonds)]
    if filt_regions:  df_view = df_view[df_view["region"].isin(filt_regions)]

    st.markdown(
        f'<div class="pipeline-hint">Selectionnez une ligne pour modifier '
        f'— <b>{len(df_view)} deal(s)</b> affiches</div>',
        unsafe_allow_html=True
    )

    df_display = df_view.copy()
    df_display["next_action_date"] = df_display["next_action_date"].apply(
        lambda d: d.isoformat() if isinstance(d, date) else ""
    )
    # Colonnes affichees
    cols_show = ["id", "nom_client", "type_client", "region", "fonds",
                 "statut", "target_aum_initial", "revised_aum", "funded_aum",
                 "raison_perte", "next_action_date", "sales_owner"]

    event = st.dataframe(
        df_display[cols_show],
        use_container_width=True,
        height=380,
        hide_index=True,
        on_select="rerun",
        selection_mode="single-row",
        column_config={
            "id":                 st.column_config.NumberColumn(
                                      "ID", width="small"),
            "nom_client":         st.column_config.TextColumn("Client"),
            "type_client":        st.column_config.TextColumn(
                                      "Type", width="small"),
            "region":             st.column_config.TextColumn(
                                      "Region", width="small"),
            "fonds":              st.column_config.TextColumn("Fonds"),
            "statut":             st.column_config.TextColumn("Statut"),
            "target_aum_initial": st.column_config.NumberColumn(
                                      "AUM Cible (EUR)", format="%.0f"),
            "revised_aum":        st.column_config.NumberColumn(
                                      "AUM Revise (EUR)", format="%.0f"),
            "funded_aum":         st.column_config.NumberColumn(
                                      "AUM Finance (EUR)", format="%.0f"),
            "raison_perte":       st.column_config.TextColumn("Raison"),
            "next_action_date":   st.column_config.TextColumn("Next Action"),
            "sales_owner":        st.column_config.TextColumn("Commercial"),
        },
        key="pipeline_ro"
    )

    selected_rows = event.selection.rows if event.selection else []

    if len(selected_rows) > 0:
        sel_row     = df_view.iloc[selected_rows[0]]
        pipeline_id = int(sel_row["id"])
        row_data    = db.get_pipeline_row_by_id(pipeline_id)

        if row_data:
            client_name    = str(row_data.get("nom_client", ""))
            current_statut = str(row_data.get("statut", "Prospect"))

            st.markdown(f"""
            <div class="detail-panel">
                <div style="font-size:0.90rem;font-weight:700;
                            color:{C_MARINE};margin-bottom:11px;">
                    Modification — <span style="color:{C_BLEU_MID};">
                    {client_name}</span>
                    &nbsp; {statut_badge(current_statut)}
                    &nbsp;<span style="font-size:0.70rem;color:#888;
                                       font-weight:400;">ID #{pipeline_id}</span>
                </div>
            </div>
            """, unsafe_allow_html=True)

            with st.container():
                with st.form(key=f"edit_{pipeline_id}"):
                    r1c1, r1c2 = st.columns(2)
                    with r1c1:
                        fi = (FONDS.index(row_data["fonds"])
                              if row_data["fonds"] in FONDS else 0)
                        new_fonds  = st.selectbox("Fonds", FONDS, index=fi)
                    with r1c2:
                        si = (STATUTS.index(current_statut)
                              if current_statut in STATUTS else 0)
                        new_statut = st.selectbox("Statut", STATUTS, index=si)

                    r2c1, r2c2, r2c3 = st.columns(3)
                    with r2c1:
                        new_target = st.number_input(
                            "AUM Cible (EUR)",
                            value=float(row_data.get(
                                "target_aum_initial", 0.0)),
                            min_value=0.0,
                            step=1_000_000.0,
                            format="%.0f")
                    with r2c2:
                        new_revised = st.number_input(
                            "AUM Revise (EUR)",
                            value=float(row_data.get("revised_aum", 0.0)),
                            min_value=0.0,
                            step=1_000_000.0,
                            format="%.0f")
                    with r2c3:
                        new_funded = st.number_input(
                            "AUM Finance (EUR)",
                            value=float(row_data.get("funded_aum", 0.0)),
                            min_value=0.0,
                            step=1_000_000.0,
                            format="%.0f")

                    r3c1, r3c2, r3c3, r3c4 = st.columns(4)
                    with r3c1:
                        ropts  = [""] + RAISONS_PERTE
                        cur_r  = str(row_data.get("raison_perte") or "")
                        ri     = ropts.index(cur_r) if cur_r in ropts else 0
                        lbl_r  = ("Raison (obligatoire)"
                                  if new_statut in ("Lost", "Paused")
                                  else "Raison Perte / Pause")
                        new_raison = st.selectbox(lbl_r, ropts, index=ri)
                    with r3c2:
                        new_concurrent = st.text_input(
                            "Concurrent",
                            value=str(
                                row_data.get("concurrent_choisi") or ""))
                    with r3c3:
                        nad = row_data.get("next_action_date")
                        if not isinstance(nad, date):
                            nad = date.today() + timedelta(days=14)
                        new_nad = st.date_input("Prochaine Action", value=nad)
                    with r3c4:
                        new_sales = st.text_input(
                            "Commercial",
                            value=str(
                                row_data.get("sales_owner") or "Non assigne"))

                    sub = st.form_submit_button(
                        "Sauvegarder les modifications",
                        use_container_width=True)

                if sub:
                    ok, msg = db.update_pipeline_row({
                        "id":                 pipeline_id,
                        "fonds":              new_fonds,
                        "statut":             new_statut,
                        "target_aum_initial": new_target,
                        "revised_aum":        new_revised,
                        "funded_aum":         new_funded,
                        "raison_perte":       new_raison,
                        "concurrent_choisi":  new_concurrent,
                        "next_action_date":   new_nad,
                        "sales_owner":        new_sales,
                    })
                    if ok:
                        st.success(
                            f"Deal {new_fonds} / {client_name} mis a jour. "
                            f"Statut : {new_statut}. "
                            f"AUM Finance : {fmt_m(new_funded)}."
                        )
                        st.rerun()
                    else:
                        st.error(msg)

            # Audit trail sobre
            df_audit = db.get_audit_log(pipeline_id)
            if not df_audit.empty:
                with st.expander(
                    f"Historique des modifications — Deal #{pipeline_id}",
                    expanded=False
                ):
                    # Rendu sobre : tableau compact sans styling excessif
                    st.dataframe(
                        df_audit,
                        use_container_width=True,
                        hide_index=True,
                        height=min(220, 46 + len(df_audit) * 36),
                        column_config={
                            "Champ":          st.column_config.TextColumn(
                                                  "Champ modifie"),
                            "Ancienne valeur": st.column_config.TextColumn(
                                                  "Ancienne"),
                            "Nouvelle valeur": st.column_config.TextColumn(
                                                  "Nouvelle"),
                            "Modifie par":    st.column_config.TextColumn(
                                                  "Par"),
                            "Date":           st.column_config.TextColumn(
                                                  "Date"),
                        }
                    )
            else:
                st.caption("Aucune modification enregistree pour ce deal.")
    else:
        st.markdown(f"""
        <div style="background:{C_MARINE}04;border:1px dashed {C_MARINE}20;
                    border-radius:6px;padding:20px;text-align:center;
                    margin-top:10px;">
            <div style="color:{C_MARINE};font-weight:600;font-size:0.86rem;">
                Selectionnez un deal dans le tableau</div>
            <div style="color:#888;font-size:0.76rem;margin-top:2px;">
                Le formulaire d'edition et l'historique s'afficheront ici</div>
        </div>""", unsafe_allow_html=True)

    # Graphique Plotly interactif — AUM par deal actif
    st.divider()
    st.markdown("#### Comparaison AUM par Deal Actif")
    df_viz = df_pipe[
        (df_pipe["target_aum_initial"] > 0) &
        (df_pipe["statut"].isin(
            ["Funded", "Soft Commit", "Due Diligence", "Initial Pitch"]))
    ].copy().head(10)

    if not df_viz.empty:
        fig_viz = go.Figure()
        clients_viz = df_viz["nom_client"].str[:18].tolist()

        for col_key, col_name, color in [
            ("target_aum_initial", "AUM Cible",   C_GRIS),
            ("revised_aum",        "AUM Revise",  C_BLEU_MID),
            ("funded_aum",         "AUM Finance", C_MARINE),
        ]:
            vals = df_viz[col_key].tolist()
            hover = [
                f"<b>{c}</b><br>{col_name} : {fmt_m(v)}"
                for c, v in zip(clients_viz, vals)
            ]
            fig_viz.add_trace(go.Bar(
                name=col_name,
                x=clients_viz,
                y=vals,
                marker_color=color,
                hovertext=hover,
                hoverinfo="text",
                marker_line_color=C_BLANC,
                marker_line_width=0.5,
            ))

        fig_viz.update_layout(
            **_plotly_layout("AUM Cible / Revise / Finance — Deals Actifs",
                             height=340),
            barmode="group",
            xaxis=dict(
                tickangle=-20, tickfont=dict(size=10, color=C_MARINE)),
            yaxis=dict(
                tickfont=dict(size=9, color=C_MARINE),
                tickformat=".2s",
            ),
            bargap=0.22,
            bargroupgap=0.06,
        )
        fig_viz.update_yaxes(
            tickvals=fig_viz.full_figure_for_development(warn=False
                ).layout.yaxis.tickvals
            if False else None,
            ticktext=None,
        )
        # Override format Y axis avec fmt_m custom via JS non disponible —
        # on utilise .2s (milliers/millions) natif Plotly
        st.plotly_chart(fig_viz, use_container_width=True,
                        config={"displayModeBar": False})

    # Tableau Lost / Paused compact
    df_lp = df_pipe[df_pipe["statut"].isin(["Lost", "Paused"])].copy()
    if not df_lp.empty:
        st.divider()
        st.markdown("#### Deals Perdus / En Pause")
        df_lp_disp = df_lp[
            ["nom_client", "fonds", "statut",
             "target_aum_initial", "raison_perte",
             "concurrent_choisi", "sales_owner"]
        ].copy()
        df_lp_disp["target_aum_initial"] = df_lp_disp[
            "target_aum_initial"].apply(fmt_m)
        st.dataframe(
            df_lp_disp,
            use_container_width=True,
            hide_index=True,
            column_config={
                "nom_client":         st.column_config.TextColumn("Client"),
                "fonds":              st.column_config.TextColumn("Fonds"),
                "statut":             st.column_config.TextColumn("Statut"),
                "target_aum_initial": st.column_config.TextColumn(
                                          "AUM Cible"),
                "raison_perte":       st.column_config.TextColumn("Raison"),
                "concurrent_choisi":  st.column_config.TextColumn("Concurrent"),
                "sales_owner":        st.column_config.TextColumn("Commercial"),
            }
        )


# ============================================================================
# ONGLET 3 : EXECUTIVE DASHBOARD — Plotly interactif + KPIs cliquables
# ============================================================================
with tab_dash:
    st.markdown('<div class="section-title">Executive Dashboard</div>',
                unsafe_allow_html=True)

    kpis = db.get_kpis()
    nb_lost_paused = kpis["nb_lost"] + kpis.get("nb_paused", 0)

    # KPI Cards cliquables — st.button stylise en KPI
    kpi_cols = st.columns(4, gap="medium")

    with kpi_cols[0]:
        st.markdown(
            kpi_card(
                "AUM Finance Total",
                fmt_m(kpis["total_funded"]),
                f"{kpis['nb_funded']} deal(s) Funded — cliquer pour detail"
            ),
            unsafe_allow_html=True
        )
        if st.button("Detail Funded", key="btn_kpi_funded",
                     use_container_width=True):
            modal_funded(_filtre_effectif)

    with kpi_cols[1]:
        st.markdown(
            kpi_card(
                "Pipeline Actif",
                fmt_m(kpis["pipeline_actif"]),
                f"{kpis['nb_deals_actifs']} deals en cours"
            ),
            unsafe_allow_html=True
        )
        if st.button("Detail Pipeline", key="btn_kpi_pipeline",
                     use_container_width=True):
            modal_pipeline_actif(_filtre_effectif)

    with kpi_cols[2]:
        st.markdown(
            kpi_card(
                "Taux Conversion",
                f"{kpis['taux_conversion']}%",
                f"{kpis['nb_funded']} funded / {kpis['nb_lost']} lost"
            ),
            unsafe_allow_html=True
        )

    with kpi_cols[3]:
        st.markdown(
            kpi_card(
                "Lost / Paused",
                str(nb_lost_paused),
                "Cliquer pour analyser"
            ),
            unsafe_allow_html=True
        )
        if nb_lost_paused > 0:
            if st.button("Detail Lost / Paused", key="btn_kpi_lost",
                         use_container_width=True):
                modal_lost(_filtre_effectif)

    st.markdown("<br>", unsafe_allow_html=True)

    # Repartition par statut — pastilles epurees
    statut_order = [s for s in STATUTS
                    if kpis["statut_repartition"].get(s, 0) > 0]
    if statut_order:
        bcols = st.columns(len(statut_order), gap="small")
        for col, s in zip(bcols, statut_order):
            c_hex = STATUT_COLORS.get(s, C_GRIS)
            with col:
                st.markdown(f"""
                <div style="background:{c_hex}16;border:1px solid {c_hex}44;
                            border-radius:5px;padding:7px;text-align:center;">
                    <div style="font-size:0.62rem;color:{C_MARINE};
                                font-weight:700;text-transform:uppercase;">
                        {s}</div>
                    <div style="font-size:1.35rem;font-weight:800;
                                color:{c_hex};">
                        {kpis['statut_repartition'][s]}</div>
                </div>""", unsafe_allow_html=True)

    st.divider()

    # Alertes actions en retard
    df_overdue = db.get_overdue_actions()
    if not df_overdue.empty:
        with st.expander(
            f"Actions en retard — {len(df_overdue)} alerte(s)",
            expanded=True
        ):
            for _, row in df_overdue.iterrows():
                nad = row.get("next_action_date")
                if isinstance(nad, date):
                    days_late = (date.today() - nad).days
                    nad_str   = nad.isoformat()
                else:
                    days_late = 0
                    nad_str   = str(nad or "—")
                owner = str(row.get("sales_owner", "")) or ""
                retard_badge = (
                    f'<span class="badge-retard">RETARD +{days_late}j</span>'
                )
                st.markdown(f"""
                <div class="alert-overdue">
                    <b>{row['nom_client']}</b> — {row['fonds']}
                    <span style="color:{C_BLEU_MID};font-weight:600;">
                        ({row['statut']})</span>
                    — Prevue le <b>{nad_str}</b>
                    &nbsp;{retard_badge}
                    {f' — <b>{owner}</b>' if owner else ''}
                </div>""", unsafe_allow_html=True)

    # Graphiques Plotly — 3 colonnes
    gcol1, gcol2, gcol3 = st.columns([1, 1, 1.2], gap="medium")

    # -- Donut : Par Type Client --
    with gcol1:
        st.markdown("#### Par Type Client")
        if kpis["aum_by_type"]:
            labels_t = list(kpis["aum_by_type"].keys())
            values_t = list(kpis["aum_by_type"].values())
            hover_t  = [fmt_m(v) for v in values_t]

            fig_type = go.Figure(go.Pie(
                labels=labels_t,
                values=values_t,
                hole=0.52,
                marker=dict(
                    colors=PLOTLY_PALETTE[:len(labels_t)],
                    line=dict(color=C_BLANC, width=2)
                ),
                hovertemplate=(
                    "<b>%{label}</b><br>"
                    "AUM Finance : %{customdata}<br>"
                    "Part : %{percent}<extra></extra>"
                ),
                customdata=hover_t,
                textinfo="percent",
                textfont=dict(size=10, color=C_BLANC),
                insidetextorientation="auto",
            ))
            fig_type.add_annotation(
                text=fmt_m(sum(values_t)),
                x=0.5, y=0.55,
                font=dict(size=11, color=C_MARINE, family="Segoe UI"),
                showarrow=False
            )
            fig_type.add_annotation(
                text="Finance",
                x=0.5, y=0.42,
                font=dict(size=8, color=C_GRIS_TXT),
                showarrow=False
            )
            fig_type.update_layout(
                **_plotly_layout("AUM par Type", height=300),
                showlegend=True,
                legend=dict(
                    orientation="v", x=1.02, y=0.5,
                    font=dict(size=9, color=C_MARINE),
                ),
                margin=dict(l=0, r=80, t=36, b=10),
            )
            st.plotly_chart(fig_type, use_container_width=True,
                            config={"displayModeBar": False})

    # -- Donut : Par Region --
    with gcol2:
        st.markdown("#### Par Region")
        aum_by_region_dash = db.get_aum_by_region()
        if aum_by_region_dash:
            labels_r = list(aum_by_region_dash.keys())
            values_r = list(aum_by_region_dash.values())
            hover_r  = [fmt_m(v) for v in values_r]

            fig_reg = go.Figure(go.Pie(
                labels=labels_r,
                values=values_r,
                hole=0.52,
                marker=dict(
                    colors=PLOTLY_PALETTE[:len(labels_r)],
                    line=dict(color=C_BLANC, width=2)
                ),
                hovertemplate=(
                    "<b>%{label}</b><br>"
                    "AUM Finance : %{customdata}<br>"
                    "Part : %{percent}<extra></extra>"
                ),
                customdata=hover_r,
                textinfo="percent",
                textfont=dict(size=10, color=C_BLANC),
                insidetextorientation="auto",
            ))
            fig_reg.add_annotation(
                text=fmt_m(sum(values_r)),
                x=0.5, y=0.55,
                font=dict(size=11, color=C_MARINE, family="Segoe UI"),
                showarrow=False
            )
            fig_reg.add_annotation(
                text="Finance",
                x=0.5, y=0.42,
                font=dict(size=8, color=C_GRIS_TXT),
                showarrow=False
            )
            fig_reg.update_layout(
                **_plotly_layout("AUM par Region", height=300),
                showlegend=True,
                legend=dict(
                    orientation="v", x=1.02, y=0.5,
                    font=dict(size=9, color=C_MARINE),
                ),
                margin=dict(l=0, r=80, t=36, b=10),
            )
            st.plotly_chart(fig_reg, use_container_width=True,
                            config={"displayModeBar": False})

    # -- Bar : AUM par Fonds --
    with gcol3:
        st.markdown("#### AUM par Fonds")
        if kpis["aum_by_fonds"]:
            fonds_sorted = sorted(
                kpis["aum_by_fonds"].items(),
                key=lambda x: x[1], reverse=True
            )
            flbls = [f[0] for f in fonds_sorted]
            fvals = [f[1] for f in fonds_sorted]
            hover_f = [fmt_m(v) for v in fvals]

            fig_fonds = go.Figure(go.Bar(
                x=fvals,
                y=flbls,
                orientation="h",
                marker=dict(
                    color=PLOTLY_PALETTE[:len(flbls)],
                    line=dict(color=C_BLANC, width=0.5)
                ),
                hovertemplate=(
                    "<b>%{y}</b><br>AUM Finance : %{customdata}"
                    "<extra></extra>"
                ),
                customdata=hover_f,
                text=hover_f,
                textposition="outside",
                textfont=dict(size=9, color=C_MARINE),
            ))
            fig_fonds.update_layout(
                **_plotly_layout("AUM Finance par Fonds", height=300),
                xaxis=dict(
                    tickfont=dict(size=8, color=C_MARINE),
                    showgrid=True,
                    gridcolor=C_GRIS,
                ),
                yaxis=dict(
                    tickfont=dict(size=9, color=C_MARINE),
                    autorange="reversed",
                ),
                margin=dict(l=10, r=60, t=36, b=10),
            )
            st.plotly_chart(fig_fonds, use_container_width=True,
                            config={"displayModeBar": False})

    st.divider()

    # Top Deals — liste visuellement epuree
    st.markdown("#### Top Deals — AUM Finance")
    df_funded_top = (
        db.get_pipeline_with_clients()
        .query("statut == 'Funded'")
        .sort_values("funded_aum", ascending=False)
        .head(10)
    )
    if not df_funded_top.empty:
        max_f = float(df_funded_top["funded_aum"].max())
        for i, (_, row) in enumerate(df_funded_top.iterrows()):
            val  = float(row["funded_aum"])
            pct  = val / max_f * 100 if max_f > 0 else 0
            rank = f"No.{i+1}"
            bar_color = C_MARINE if i == 0 else C_BLEU_MID if i < 3 else C_BLEU_PAL
            st.markdown(f"""
            <div style="display:flex;align-items:center;gap:10px;
                        margin:4px 0;padding:7px 12px;background:#F7FAFD;
                        border-radius:5px;border:1px solid {C_MARINE}0A;">
                <div style="font-size:0.74rem;font-weight:700;
                            color:{C_BLEU_MID};min-width:28px;">{rank}</div>
                <div style="flex:1;">
                    <div style="font-size:0.81rem;font-weight:700;
                                color:{C_MARINE};">{row['nom_client']}</div>
                    <div style="font-size:0.68rem;color:#777;">
                        {row['fonds']} &middot; {row['type_client']}
                        &middot; {row['region']}</div>
                    <div style="background:{C_GRIS};border-radius:2px;
                                height:3px;margin-top:4px;overflow:hidden;">
                        <div style="background:{bar_color};width:{pct:.0f}%;
                                    height:100%;border-radius:2px;"></div>
                    </div>
                </div>
                <div style="text-align:right;min-width:75px;">
                    <div style="font-size:0.88rem;font-weight:800;
                                color:{C_MARINE};">{fmt_m(val)}</div>
                    <div style="font-size:0.62rem;color:#999;">finance</div>
                </div>
            </div>""", unsafe_allow_html=True)


# ============================================================================
# ONGLET 4 : SALES TRACKING — Charge par commercial + Next Actions
# ============================================================================
with tab_sales:
    st.markdown(
        '<div class="section-title">Sales Tracking — Suivi par Commercial'
        '</div>',
        unsafe_allow_html=True
    )

    df_sm = db.get_sales_metrics()
    df_na = db.get_next_actions_by_sales(days_ahead=30)

    if df_sm.empty:
        st.info("Aucune donnee de pipeline disponible.")
    else:
        # Cards commerciaux
        n_own  = len(df_sm)
        n_cols = min(n_own, 3)
        s_cols = st.columns(n_cols, gap="medium")

        for i, (_, row) in enumerate(df_sm.iterrows()):
            retard_val = int(row.get("Actions en retard", 0))
            if retard_val > 0:
                retard_html = (
                    f'<span class="badge-retard">'
                    f'RETARD : {retard_val}</span>'
                )
            else:
                retard_html = (
                    f'<span style="color:{C_BLEU_MID};font-size:0.75rem;'
                    f'font-weight:600;">A jour</span>'
                )

            with s_cols[i % n_cols]:
                st.markdown(f"""
                <div class="sales-card">
                    <div class="sales-card-name">{row['Commercial']}</div>
                    <div style="display:grid;grid-template-columns:1fr 1fr;
                                gap:8px;">
                        <div>
                            <div class="sales-metric">Total Deals</div>
                            <div class="sales-metric-val">
                                {int(row['Nb Deals'])}</div>
                        </div>
                        <div>
                            <div class="sales-metric">Funded</div>
                            <div class="sales-metric-val">
                                {int(row['Funded'])}</div>
                        </div>
                        <div>
                            <div class="sales-metric">AUM Finance</div>
                            <div class="sales-metric-acc">
                                {fmt_m(float(row['AUM Finance']))}</div>
                        </div>
                        <div>
                            <div class="sales-metric">Pipeline Actif</div>
                            <div class="sales-metric-val">
                                {fmt_m(float(row['Pipeline Actif']))}</div>
                        </div>
                        <div>
                            <div class="sales-metric">Actifs / Perdus</div>
                            <div class="sales-metric-val">
                                {int(row['Actifs'])} / {int(row['Perdus'])}</div>
                        </div>
                        <div>
                            <div class="sales-metric">Alertes</div>
                            <div style="margin-top:2px;">{retard_html}</div>
                        </div>
                    </div>
                </div><br>""", unsafe_allow_html=True)

        st.divider()

        # Graphique Plotly — AUM par commercial
        st.markdown("#### AUM Finance vs Pipeline par Commercial")
        if df_sm["AUM Finance"].sum() > 0:
            owners_l = df_sm["Commercial"].tolist()
            aum_v    = df_sm["AUM Finance"].tolist()
            pipe_v   = df_sm["Pipeline Actif"].tolist()
            hover_aum  = [fmt_m(v) for v in aum_v]
            hover_pipe = [fmt_m(v) for v in pipe_v]

            fig_sales = go.Figure()
            fig_sales.add_trace(go.Bar(
                name="AUM Finance",
                x=owners_l,
                y=aum_v,
                marker_color=C_MARINE,
                hovertemplate=(
                    "<b>%{x}</b><br>AUM Finance : %{customdata}"
                    "<extra></extra>"
                ),
                customdata=hover_aum,
                marker_line_color=C_BLANC,
                marker_line_width=0.5,
            ))
            fig_sales.add_trace(go.Bar(
                name="Pipeline Actif",
                x=owners_l,
                y=pipe_v,
                marker_color=C_BLEU_MID,
                hovertemplate=(
                    "<b>%{x}</b><br>Pipeline Actif : %{customdata}"
                    "<extra></extra>"
                ),
                customdata=hover_pipe,
                marker_line_color=C_BLANC,
                marker_line_width=0.5,
            ))
            fig_sales.update_layout(
                **_plotly_layout("", height=320),
                barmode="group",
                xaxis=dict(tickfont=dict(size=10, color=C_MARINE)),
                yaxis=dict(
                    tickfont=dict(size=9, color=C_MARINE),
                    showgrid=True, gridcolor=C_GRIS,
                ),
                bargap=0.25,
                bargroupgap=0.08,
            )
            st.plotly_chart(fig_sales, use_container_width=True,
                            config={"displayModeBar": False})

        st.divider()

        # Tableau recapitulatif commercial
        st.markdown("#### Tableau de Bord Commercial")
        df_sd = df_sm.copy()
        df_sd["AUM Finance"]   = df_sd["AUM Finance"].apply(fmt_m)
        df_sd["Pipeline Actif"]= df_sd["Pipeline Actif"].apply(fmt_m)
        st.dataframe(
            df_sd,
            use_container_width=True,
            hide_index=True,
            height=min(320, 50 + len(df_sd) * 40)
        )

        st.divider()

        # Prochaines Actions — priorisees, filtrees par commercial
        st.markdown("#### Prochaines Actions — 30 jours")
        if df_na.empty:
            st.info("Aucune action planifiee dans les 30 prochains jours.")
        else:
            owners_na = ["Tous"] + sorted(
                df_na["sales_owner"].dropna().unique().tolist()
            )
            filter_o = st.selectbox(
                "Filtrer par commercial", owners_na, key="sel_na_owner")
            df_nav = (df_na if filter_o == "Tous"
                      else df_na[df_na["sales_owner"] == filter_o])

            today_d = date.today()
            for _, row in df_nav.iterrows():
                nad = row.get("next_action_date")
                if isinstance(nad, date):
                    days_diff = (nad - today_d).days
                    nad_str   = nad.isoformat()
                    if days_diff <= 3:
                        urgence = (
                            f'<span class="badge-retard">URGENT +{days_diff}j'
                            f'</span>'
                        )
                    elif days_diff <= 7:
                        urgence = (
                            f'<span style="background:{C_MARINE};'
                            f'color:{C_BLANC};border-radius:3px;'
                            f'padding:1px 6px;font-size:0.67rem;'
                            f'font-weight:700;">J+{days_diff}</span>'
                        )
                    else:
                        urgence = (
                            f'<span style="background:{C_GRIS};'
                            f'color:{C_MARINE};border-radius:3px;'
                            f'padding:1px 6px;font-size:0.67rem;'
                            f'font-weight:600;">J+{days_diff}</span>'
                        )
                else:
                    nad_str = "—"
                    urgence = ""

                aum_str = fmt_m(float(row.get("revised_aum", 0) or 0))
                st.markdown(f"""
                <div style="display:flex;align-items:center;gap:10px;
                            padding:6px 12px;margin:3px 0;
                            background:#F7FAFD;border-radius:4px;
                            border:1px solid {C_MARINE}0A;
                            font-size:0.79rem;">
                    <div style="flex:1;">
                        <span style="font-weight:700;color:{C_MARINE};">
                            {row.get('nom_client','')}</span>
                        <span style="color:#777;"> — {row.get('fonds','')}
                        </span>
                        <span style="color:{C_BLEU_MID};font-size:0.72rem;">
                            ({row.get('statut','')})</span>
                    </div>
                    <div style="color:#555;min-width:80px;font-size:0.75rem;">
                        {nad_str}</div>
                    <div style="min-width:60px;">{urgence}</div>
                    <div style="color:{C_BLEU_MID};font-weight:700;
                                min-width:75px;text-align:right;">
                        {aum_str}</div>
                    <div style="color:#888;font-size:0.72rem;min-width:80px;">
                        {row.get('sales_owner','')}</div>
                </div>""", unsafe_allow_html=True)


# ============================================================================
# ONGLET 5 : PERFORMANCE ET NAV — Plotly + tableau performances
# ============================================================================
with tab_perf:
    st.markdown(
        '<div class="section-title">Performance et NAV — Module Analytique'
        '</div>',
        unsafe_allow_html=True
    )

    col_upload, col_demo = st.columns([2, 1], gap="medium")
    with col_upload:
        nav_file = st.file_uploader(
            "Charger un fichier NAV (Excel ou CSV)",
            type=["xlsx", "xls", "csv"],
            help="Colonnes attendues : Date, Fonds, NAV"
        )
    with col_demo:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Generer un fichier NAV de demonstration",
                     use_container_width=True):
            today_ts = pd.Timestamp(date.today())
            dates_d  = pd.date_range(
                end=today_ts, periods=252, freq="B")
            demo_rows = []
            for fonds_d, seed in zip(FONDS, range(len(FONDS))):
                rng   = np.random.default_rng(seed=seed + 42)
                ret   = rng.normal(0.0003, 0.008, size=len(dates_d))
                nav_d = 100.0 * np.cumprod(1 + ret)
                for d, n in zip(dates_d, nav_d):
                    demo_rows.append({
                        "Date": d.strftime("%Y-%m-%d"),
                        "Fonds": fonds_d,
                        "NAV": round(float(n), 6)
                    })
            df_demo = pd.DataFrame(demo_rows)
            buf_demo = io.BytesIO()
            df_demo.to_excel(buf_demo, index=False, engine="openpyxl")
            st.download_button(
                "Telecharger le fichier de demonstration",
                data=buf_demo.getvalue(),
                file_name="nav_demo.xlsx",
                mime="application/vnd.openxmlformats-officedocument"
                     ".spreadsheetml.sheet",
                use_container_width=True,
            )

    if nav_file is not None:
        try:
            df_nav_raw = (
                pd.read_csv(nav_file)
                if nav_file.name.endswith(".csv")
                else pd.read_excel(nav_file)
            )

            # Normalisation colonnes
            df_nav_raw.columns = [str(c).strip() for c in df_nav_raw.columns]
            col_map = {}
            for c in df_nav_raw.columns:
                c_low = c.lower()
                if "date" in c_low:           col_map[c] = "Date"
                elif "fonds" in c_low or "fund" in c_low: col_map[c] = "Fonds"
                elif "nav" in c_low:          col_map[c] = "NAV"
            df_nav_raw = df_nav_raw.rename(columns=col_map)

            required = {"Date", "Fonds", "NAV"}
            missing  = required - set(df_nav_raw.columns)
            if missing:
                st.error(
                    f"Colonnes manquantes : {', '.join(missing)}. "
                    f"Colonnes detectees : {list(df_nav_raw.columns)}"
                )
            else:
                df_nav_raw["Date"] = pd.to_datetime(
                    df_nav_raw["Date"], errors="coerce")
                df_nav_raw["NAV"]  = pd.to_numeric(
                    df_nav_raw["NAV"], errors="coerce")
                df_nav_raw = df_nav_raw.dropna(
                    subset=["Date", "NAV"]).copy()

                # Filtres interactifs
                pf1, pf2, pf3 = st.columns(3)
                fonds_nav_opts = sorted(df_nav_raw["Fonds"].unique().tolist())
                with pf1:
                    fonds_sel_nav = st.multiselect(
                        "Fonds", fonds_nav_opts, default=fonds_nav_opts,
                        key="nav_fonds_sel")
                with pf2:
                    d_debut = st.date_input(
                        "Date debut",
                        value=df_nav_raw["Date"].min().date(),
                        key="nav_d_debut")
                with pf3:
                    d_fin = st.date_input(
                        "Date fin",
                        value=df_nav_raw["Date"].max().date(),
                        key="nav_d_fin")

                mask = (
                    df_nav_raw["Fonds"].isin(fonds_sel_nav) &
                    (df_nav_raw["Date"] >= pd.Timestamp(d_debut)) &
                    (df_nav_raw["Date"] <= pd.Timestamp(d_fin))
                )
                df_nav_f = df_nav_raw[mask].copy()

                if df_nav_f.empty:
                    st.warning("Aucune donnee pour ces filtres.")
                else:
                    # Pivot pour Base 100
                    pivot = df_nav_f.pivot_table(
                        index="Date", columns="Fonds",
                        values="NAV", aggfunc="last"
                    ).sort_index()

                    base100 = pivot.copy()
                    for col in base100.columns:
                        first = base100[col].dropna()
                        if not first.empty and float(first.iloc[0]) != 0:
                            base100[col] = (
                                base100[col] / float(first.iloc[0]) * 100
                            )

                    # Stockage pour export PDF
                    st.session_state["perf_data"]   = None
                    st.session_state["nav_base100"]  = base100

                    # --- Graphique Base 100 — Plotly interactif ---
                    st.markdown("#### Courbe NAV — Base 100")
                    fig_nav = go.Figure()

                    palette_nav = [
                        C_MARINE, C_BLEU_MID, C_BLEU_PAL, C_BLEU_DEP,
                        "#2C7FB8", "#004F8C", "#6BAED6", "#08519C",
                    ]
                    dash_styles = ["solid", "dash", "dot",
                                   "dashdot", "longdash"]

                    for i, fonds_n in enumerate(base100.columns):
                        series_n = base100[fonds_n].dropna()
                        if series_n.empty:
                            continue
                        color_n = palette_nav[i % len(palette_nav)]
                        dash_n  = dash_styles[i % len(dash_styles)]

                        fig_nav.add_trace(go.Scatter(
                            x=series_n.index,
                            y=series_n.values,
                            name=fonds_n,
                            mode="lines",
                            line=dict(
                                color=color_n, width=1.8,
                                dash=dash_n
                            ),
                            hovertemplate=(
                                f"<b>{fonds_n}</b><br>"
                                "Date : %{x|%d/%m/%Y}<br>"
                                "Base 100 : %{y:.2f}<extra></extra>"
                            ),
                        ))

                    fig_nav.add_hline(
                        y=100,
                        line=dict(color=C_GRIS, width=0.8, dash="dot"),
                        annotation_text="Base 100",
                        annotation_font=dict(size=8, color=C_GRIS_TXT),
                    )
                    fig_nav.update_layout(
                        **_plotly_layout(
                            f"Performance NAV — Base 100 — "
                            f"{d_debut.strftime('%d/%m/%Y')} au "
                            f"{d_fin.strftime('%d/%m/%Y')}",
                            height=380
                        ),
                        xaxis=dict(
                            tickfont=dict(size=9, color=C_MARINE),
                            showgrid=False,
                        ),
                        yaxis=dict(
                            tickfont=dict(size=9, color=C_MARINE),
                            showgrid=True, gridcolor=C_GRIS,
                            zeroline=False,
                        ),
                        hovermode="x unified",
                    )
                    st.plotly_chart(fig_nav, use_container_width=True,
                                    config={"displayModeBar": True,
                                            "modeBarButtonsToRemove":
                                                ["lasso2d", "select2d"]})

                    # --- Calcul des performances ---
                    today_ts  = pd.Timestamp(date.today())
                    one_m_ago = today_ts - pd.DateOffset(months=1)
                    jan_1     = pd.Timestamp(f"{date.today().year}-01-01")

                    perf_rows = []
                    for fonds_n in pivot.columns:
                        series_n = pivot[fonds_n].dropna()
                        if series_n.empty:
                            continue
                        nav_last  = float(series_n.iloc[-1])
                        nav_first = float(series_n.iloc[0])

                        s_1m = series_n[series_n.index >= one_m_ago]
                        p1m  = (
                            (nav_last / float(s_1m.iloc[0]) - 1) * 100
                            if len(s_1m) > 0 and float(s_1m.iloc[0]) != 0
                            else float("nan")
                        )

                        s_ytd = series_n[series_n.index >= jan_1]
                        pytd  = (
                            (nav_last / float(s_ytd.iloc[0]) - 1) * 100
                            if len(s_ytd) > 0 and float(s_ytd.iloc[0]) != 0
                            else float("nan")
                        )

                        pp = (
                            (nav_last / nav_first - 1) * 100
                            if nav_first != 0 else float("nan")
                        )

                        b100s = base100[fonds_n].dropna()
                        nb100 = (float(b100s.iloc[-1])
                                 if not b100s.empty else float("nan"))

                        perf_rows.append({
                            "Fonds":            fonds_n,
                            "NAV Derniere":     round(nav_last, 4),
                            "Base 100 Actuel":  (round(nb100, 2)
                                                  if not np.isnan(nb100)
                                                  else None),
                            "Perf 1M (%)":      (round(p1m, 2)
                                                  if not np.isnan(p1m)
                                                  else None),
                            "Perf YTD (%)":     (round(pytd, 2)
                                                  if not np.isnan(pytd)
                                                  else None),
                            "Perf Periode (%)": (round(pp, 2)
                                                  if not np.isnan(pp)
                                                  else None),
                        })

                    if perf_rows:
                        df_pt = pd.DataFrame(perf_rows)
                        # Stockage pour PDF
                        st.session_state["perf_data"] = df_pt

                        # Tableau HTML des performances
                        st.markdown("#### Tableau des Performances")

                        def _fp(val):
                            if val is None or (
                                isinstance(val, float) and np.isnan(val)
                            ):
                                return (f'<span style="color:#999;">'
                                        f'n.d.</span>')
                            c_perf = "#1A7A3C" if val >= 0 else "#8B2020"
                            s = "+" if val > 0 else ""
                            return (
                                f'<span style="color:{c_perf};'
                                f'font-weight:700;">{s}{val:.2f}%</span>'
                            )

                        tbl_h = f"""
                        <table style="width:100%;border-collapse:collapse;
                                      font-size:0.80rem;">
                          <thead>
                            <tr style="background:{C_MARINE};color:white;">
                              <th style="padding:7px 11px;text-align:left;">
                                  Fonds</th>
                              <th style="padding:7px 11px;text-align:right;">
                                  NAV</th>
                              <th style="padding:7px 11px;text-align:right;">
                                  Base 100</th>
                              <th style="padding:7px 11px;text-align:right;">
                                  Perf 1M</th>
                              <th style="padding:7px 11px;text-align:right;">
                                  Perf YTD</th>
                              <th style="padding:7px 11px;text-align:right;">
                                  Perf Periode</th>
                            </tr>
                          </thead><tbody>"""
                        for ii, r in enumerate(perf_rows):
                            bg = "#F7FAFD" if ii % 2 == 0 else C_BLANC
                            b100_disp = (
                                f"{r['Base 100 Actuel']:.2f}"
                                if r['Base 100 Actuel'] is not None
                                else "n.d."
                            )
                            tbl_h += f"""
                            <tr style="background:{bg};
                                       border-bottom:1px solid {C_GRIS};">
                              <td style="padding:6px 11px;font-weight:600;
                                         color:{C_MARINE};">
                                  {r['Fonds']}</td>
                              <td style="padding:6px 11px;text-align:right;
                                         color:{C_MARINE};">
                                  {r['NAV Derniere']:.4f}</td>
                              <td style="padding:6px 11px;text-align:right;
                                         color:{C_MARINE};">
                                  {b100_disp}</td>
                              <td style="padding:6px 11px;text-align:right;">
                                  {_fp(r['Perf 1M (%)'])}</td>
                              <td style="padding:6px 11px;text-align:right;">
                                  {_fp(r['Perf YTD (%)'])}</td>
                              <td style="padding:6px 11px;text-align:right;">
                                  {_fp(r['Perf Periode (%)'])}</td>
                            </tr>"""
                        tbl_h += "</tbody></table>"
                        st.markdown(tbl_h, unsafe_allow_html=True)
                        st.markdown("<br>", unsafe_allow_html=True)

                        st.download_button(
                            "Exporter les performances en CSV",
                            data=df_pt.to_csv(index=False).encode("utf-8"),
                            file_name=(
                                f"performances_{date.today().isoformat()}"
                                f".csv"),
                            mime="text/csv",
                        )

                        # Graphique YTD — Plotly interactif
                        ytd_data = [r for r in perf_rows
                                    if r["Perf YTD (%)"] is not None]
                        if ytd_data:
                            ytd_data.sort(
                                key=lambda r: r["Perf YTD (%)"],
                                reverse=True
                            )
                            st.markdown(
                                "#### Comparaison Performances YTD")

                            fonds_y = [r["Fonds"] for r in ytd_data]
                            vals_y  = [r["Perf YTD (%)"] for r in ytd_data]
                            bar_cols_y = [
                                C_MARINE if v >= 0 else "#AAAAAA"
                                for v in vals_y
                            ]
                            hover_y = [
                                (f"<b>{f}</b><br>"
                                 f"Perf YTD : {'+' if v >= 0 else ''}"
                                 f"{v:.2f}%<extra></extra>")
                                for f, v in zip(fonds_y, vals_y)
                            ]

                            fig_ytd = go.Figure(go.Bar(
                                x=fonds_y,
                                y=vals_y,
                                marker=dict(
                                    color=bar_cols_y,
                                    line=dict(color=C_BLANC, width=0.5)
                                ),
                                hovertemplate=hover_y,
                                text=[
                                    f"{'+' if v >= 0 else ''}{v:.2f}%"
                                    for v in vals_y
                                ],
                                textposition="outside",
                                textfont=dict(size=9, color=C_MARINE),
                            ))
                            fig_ytd.add_hline(
                                y=0,
                                line=dict(
                                    color=C_MARINE, width=0.8),
                            )
                            fig_ytd.update_layout(
                                **_plotly_layout(
                                    "Performance YTD par Fonds (%)",
                                    height=320
                                ),
                                xaxis=dict(
                                    tickfont=dict(size=10, color=C_MARINE),
                                    tickangle=-15,
                                ),
                                yaxis=dict(
                                    tickfont=dict(size=9, color=C_MARINE),
                                    showgrid=True,
                                    gridcolor=C_GRIS,
                                    zeroline=False,
                                    ticksuffix="%",
                                ),
                                showlegend=False,
                            )
                            st.plotly_chart(
                                fig_ytd,
                                use_container_width=True,
                                config={"displayModeBar": False}
                            )

        except Exception as e:
            st.error(f"Erreur de traitement du fichier NAV : {e}")
            import traceback
            with st.expander("Details de l'erreur"):
                st.code(traceback.format_exc())

    else:
        st.markdown(f"""
        <div style="background:{C_MARINE}04;border:2px dashed {C_MARINE}1A;
                    border-radius:8px;padding:44px;text-align:center;
                    margin-top:12px;">
            <div style="font-size:0.94rem;font-weight:700;color:{C_MARINE};
                        margin-bottom:5px;">
                Module Performance et NAV</div>
            <div style="color:#777;font-size:0.80rem;max-width:380px;
                        margin:0 auto;line-height:1.65;">
                Chargez un fichier Excel ou CSV avec les colonnes
                <code>Date</code>, <code>Fonds</code>, <code>NAV</code>
                pour generer les courbes Base 100 et le tableau de
                performances.<br><br>
                Utilisez le bouton <b>Generer un fichier NAV de demonstration</b>
                pour un exemple testable immediatement.
            </div>
        </div>""", unsafe_allow_html=True)
