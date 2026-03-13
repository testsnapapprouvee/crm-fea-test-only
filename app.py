# =============================================================================
# app.py — CRM & Reporting Tool — Asset Management — Edition Enterprise
# Charte Amundi : Marine #002D54 | Ciel #00A8E1
# Staff Engineer refactoring :
#   - Zero emoji
#   - Master-Detail avec audit trail
#   - Onglet Sales Tracking
#   - Module Performance NAV securise
#   - Export PDF avec perf_data
# Lancement : streamlit run app.py
# =============================================================================

import streamlit as st
import pandas as pd
import numpy as np
from datetime import date, timedelta
import io
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))

import database as db
import pdf_generator as pdf_gen

# ---------------------------------------------------------------------------
# CONSTANTES CHARTE
# ---------------------------------------------------------------------------
C_MARINE    = "#002D54"
C_CIEL      = "#00A8E1"
C_GRIS      = "#E0E0E0"
C_BLANC     = "#FFFFFF"
C_BLEU_MID  = "#1A6B9A"
C_BLEU_PALE = "#7BC8E8"
C_BLEU_DEEP = "#003F7A"

TYPES_CLIENT      = ["IFA", "Wholesale", "Instit", "Family Office"]
REGIONS           = ["GCC", "EMEA", "APAC", "Nordics",
                     "Asia ex-Japan", "North America", "LatAm"]
FONDS             = ["Global Value", "International Fund", "Income Builder",
                     "Resilient Equity", "Private Debt", "Active ETFs"]
STATUTS           = ["Prospect", "Initial Pitch", "Due Diligence",
                     "Soft Commit", "Funded", "Paused", "Lost", "Redeemed"]
STATUTS_ACTIFS    = ["Prospect", "Initial Pitch", "Due Diligence", "Soft Commit"]
RAISONS_PERTE     = ["Pricing", "Track Record", "Macro", "Competitor", "Autre"]
TYPES_INTERACTION = ["Call", "Meeting", "Email", "Roadshow", "Conference", "Autre"]
NAV_PALETTE       = [C_CIEL, C_MARINE, C_BLEU_MID, C_BLEU_PALE, C_BLEU_DEEP,
                     "#5BA3C9", "#2C8FBF", "#004F8C", "#A8D8EE", "#003060"]


# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="CRM Asset Management",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ---------------------------------------------------------------------------
# CSS — CHARTE AMUNDI STRICTE, ZERO EMOJI
# ---------------------------------------------------------------------------
st.markdown(f"""
<style>
    .stApp, .main .block-container {{
        background-color: {C_BLANC};
        color: {C_MARINE};
        font-family: 'Segoe UI', Arial, sans-serif;
    }}
    [data-testid="stSidebar"] {{ background-color: {C_MARINE}; }}
    [data-testid="stSidebar"] * {{ color: {C_BLANC} !important; }}
    [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3 {{ color: {C_CIEL} !important; }}

    .crm-header {{
        background: linear-gradient(135deg, {C_MARINE} 0%, #003D6B 100%);
        color: {C_BLANC};
        padding: 20px 26px;
        border-radius: 8px;
        margin-bottom: 18px;
        border-left: 5px solid {C_CIEL};
    }}
    .crm-header h1 {{
        color: {C_BLANC} !important; margin: 0;
        font-size: 1.6rem; font-weight: 700; letter-spacing: -0.2px;
    }}
    .crm-header p {{
        color: {C_CIEL}; margin: 3px 0 0 0; font-size: 0.84rem;
    }}

    .kpi-card {{
        background: linear-gradient(135deg, {C_MARINE} 0%, #004070 100%);
        padding: 16px 12px; border-radius: 8px; text-align: center;
        border: 1px solid {C_CIEL}33;
    }}
    .kpi-label {{
        font-size: 0.7rem; color: {C_CIEL}; text-transform: uppercase;
        letter-spacing: 0.8px; margin-bottom: 7px; font-weight: 600;
    }}
    .kpi-value {{ font-size: 1.55rem; font-weight: 800; color: {C_BLANC}; }}
    .kpi-sub   {{ font-size: 0.68rem; color: {C_GRIS}; margin-top: 3px; }}

    .section-title {{
        font-size: 1.0rem; font-weight: 700; color: {C_MARINE};
        border-bottom: 3px solid {C_CIEL}; padding-bottom: 5px;
        margin: 16px 0 11px 0;
    }}

    /* Master-Detail */
    .detail-panel {{
        background: linear-gradient(135deg, #F0F6FC 0%, #E8F2FA 100%);
        border: 1.5px solid {C_CIEL}55; border-radius: 10px;
        padding: 20px 22px 16px 22px; margin-top: 16px;
        box-shadow: 0 3px 16px {C_MARINE}0E;
    }}

    /* Pipeline hint */
    .pipeline-hint {{
        background: {C_CIEL}0F; border-left: 3px solid {C_CIEL};
        border-radius: 0 5px 5px 0; padding: 7px 13px;
        font-size: 0.81rem; color: {C_MARINE}; margin-bottom: 10px;
    }}

    /* Alertes overdue */
    .alert-overdue {{
        background: #FFF8E1; border-left: 4px solid {C_CIEL};
        border-radius: 0 5px 5px 0; padding: 8px 13px; margin: 4px 0;
        font-size: 0.81rem; color: {C_MARINE};
    }}

    /* Onglets */
    .stTabs [data-baseweb="tab-list"] {{
        background: {C_BLANC}; border-bottom: 2px solid {C_MARINE}18; gap: 3px;
    }}
    .stTabs [data-baseweb="tab"] {{
        color: {C_MARINE}; font-weight: 600; font-size: 0.84rem;
        padding: 9px 18px; border-radius: 5px 5px 0 0;
        background: {C_MARINE}08;
    }}
    .stTabs [aria-selected="true"] {{
        background: {C_MARINE} !important; color: {C_BLANC} !important;
    }}

    /* Boutons */
    .stButton > button {{
        background: {C_MARINE}; color: {C_BLANC}; border: none;
        border-radius: 5px; font-weight: 600; padding: 7px 18px;
        font-size: 0.85rem; transition: all 0.18s;
    }}
    .stButton > button:hover {{
        background: {C_CIEL}; color: {C_BLANC};
        box-shadow: 0 3px 10px {C_CIEL}44;
    }}

    /* Labels */
    .stSelectbox label, .stTextInput label, .stNumberInput label,
    .stDateInput label, .stTextArea label, .stRadio label {{
        color: {C_MARINE} !important; font-weight: 600; font-size: 0.8rem;
    }}

    /* Sales tracking */
    .sales-card {{
        background: {C_BLANC}; border: 1px solid {C_MARINE}18;
        border-radius: 8px; padding: 16px; border-top: 3px solid {C_CIEL};
    }}
    .sales-card-name {{
        font-size: 0.95rem; font-weight: 700; color: {C_MARINE};
        margin-bottom: 10px; padding-bottom: 6px;
        border-bottom: 1px solid {C_GRIS};
    }}
    .sales-metric {{
        font-size: 0.75rem; color: #666; margin-bottom: 2px;
    }}
    .sales-metric-val {{
        font-size: 1.05rem; font-weight: 700; color: {C_MARINE};
    }}
    .sales-metric-ciel {{
        font-size: 1.05rem; font-weight: 700; color: {C_CIEL};
    }}

    /* Perf table */
    .perf-positive {{ color: {C_CIEL}; font-weight: 700; }}
    .perf-negative {{ color: #8B2020; font-weight: 700; }}

    h1, h2, h3, h4 {{ color: {C_MARINE} !important; }}
    hr {{ border-color: {C_MARINE}15; }}
    code {{ background: {C_MARINE}0A; color: {C_MARINE}; border-radius: 3px; }}
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

def fmt_m(v: float) -> str:
    if v >= 1_000_000_000: return f"${v/1_000_000_000:.2f}Md"
    if v >= 1_000_000:     return f"${v/1_000_000:.1f}M"
    if v >= 1_000:         return f"${v/1_000:.0f}k"
    return f"${v:.0f}"


def kpi_card(label: str, value: str, sub: str = "") -> str:
    return (f'<div class="kpi-card"><div class="kpi-label">{label}</div>'
            f'<div class="kpi-value">{value}</div>'
            + (f'<div class="kpi-sub">{sub}</div>' if sub else "")
            + "</div>")


def statut_badge_html(statut: str) -> str:
    colors = {
        "Funded":        (C_CIEL,    C_BLANC),
        "Soft Commit":   (C_BLEU_MID, C_BLANC),
        "Due Diligence": ("#004F8C",  C_BLANC),
        "Initial Pitch": ("#3A7EBA",  C_BLANC),
        "Prospect":      (f"{C_MARINE}22", C_MARINE),
        "Lost":          (C_GRIS,    "#555"),
        "Paused":        ("#C8D8E8",  C_MARINE),
        "Redeemed":      ("#D0E4F0",  C_MARINE),
    }
    bg, fg = colors.get(statut, (C_GRIS, "#555"))
    return (f'<span style="padding:2px 10px;border-radius:10px;'
            f'font-size:0.76rem;font-weight:600;'
            f'background:{bg};color:{fg};">{statut}</span>')


# ---------------------------------------------------------------------------
# SIDEBAR
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown(f"""
    <div style="text-align:center;padding:8px 0 18px 0;">
        <div style="font-size:1.0rem;font-weight:800;color:{C_CIEL};letter-spacing:0.5px;">
            CRM Asset Management
        </div>
        <div style="font-size:0.7rem;color:{C_GRIS};margin-top:3px;">
            {date.today().strftime("%d %B %Y")}
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.divider()

    kpis_side = db.get_kpis()
    st.markdown(f'<div style="color:{C_CIEL};font-size:0.72rem;font-weight:700;'
                f'text-transform:uppercase;letter-spacing:1px;margin-bottom:7px;">'
                f'Apercu</div>', unsafe_allow_html=True)

    for label, val in [
        ("AUM Finance Total", fmt_m(kpis_side["total_funded"])),
        ("Pipeline Actif",    fmt_m(kpis_side["pipeline_actif"])),
    ]:
        st.markdown(f"""
        <div style="background:{C_BLANC}15;padding:10px;border-radius:7px;margin-bottom:6px;">
            <div style="font-size:0.68rem;color:{C_GRIS};">{label}</div>
            <div style="font-size:1.2rem;font-weight:800;color:{C_CIEL};">{val}</div>
        </div>""", unsafe_allow_html=True)

    st.divider()
    st.markdown(f'<div style="color:{C_CIEL};font-size:0.72rem;font-weight:700;'
                f'text-transform:uppercase;letter-spacing:1px;margin-bottom:7px;">'
                f'Export PDF</div>', unsafe_allow_html=True)

    mode_comex = st.toggle(
        "Mode Comex — Anonymisation",
        value=False,
        help="Remplace les noms clients par Type-Region dans le PDF."
    )
    if mode_comex:
        st.caption("Anonymisation activee : aucun nom client dans l'export.")

    # Recuperation des donnees perf depuis session_state (si disponibles)
    perf_data_pdf    = st.session_state.get("perf_data", None)
    nav_base100_pdf  = st.session_state.get("nav_base100", None)

    if perf_data_pdf is not None:
        st.caption(f"Donnees NAV disponibles — {len(perf_data_pdf)} fonds.")

    if st.button("Generer le rapport PDF", use_container_width=True):
        with st.spinner("Generation en cours..."):
            try:
                pdf_bytes = pdf_gen.generate_pdf(
                    pipeline_df   = db.get_pipeline_with_clients(),
                    kpis          = db.get_kpis(),
                    mode_comex    = mode_comex,
                    perf_data     = perf_data_pdf,
                    nav_base100_df= nav_base100_pdf,
                )
                fname = (f"report_comex_{date.today().isoformat()}.pdf"
                         if mode_comex
                         else f"report_{date.today().isoformat()}.pdf")
                st.download_button(
                    "Telecharger le rapport",
                    data=pdf_bytes, file_name=fname,
                    mime="application/pdf",
                    use_container_width=True
                )
                st.success("Rapport genere.")
            except Exception as e:
                st.error(f"Erreur de generation : {e}")

    st.divider()
    st.caption("Version 3.0 — Enterprise Edition")


# ---------------------------------------------------------------------------
# EN-TETE
# ---------------------------------------------------------------------------
st.markdown(f"""
<div class="crm-header">
    <h1>CRM &amp; Reporting — Asset Management</h1>
    <p>Gestion du pipeline commercial &middot; Suivi des mandats &middot;
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
            with st.form("form_add_client", clear_on_submit=True):
                c1, c2 = st.columns(2)
                with c1:
                    nom_client  = st.text_input("Nom du Client *")
                    type_client = st.selectbox("Type Client *", TYPES_CLIENT)
                with c2:
                    region = st.selectbox("Region *", REGIONS)
                sub_c = st.form_submit_button("Enregistrer", use_container_width=True)
            if sub_c:
                if not nom_client.strip():
                    st.error("Nom du client obligatoire.")
                else:
                    try:
                        db.add_client(nom_client.strip(), type_client, region)
                        st.success(f"Client {nom_client} ajoute.")
                    except Exception as e:
                        st.warning("Ce client existe deja."
                                   if "UNIQUE" in str(e) else f"Erreur : {e}")

        with st.expander("Ajouter un Deal Pipeline", expanded=True):
            clients_dict = db.get_client_options()
            sales_owners = db.get_sales_owners() or ["Non assigne"]

            if not clients_dict:
                st.info("Ajoutez d'abord un client.")
            else:
                with st.form("form_add_deal", clear_on_submit=True):
                    ca, cb = st.columns(2)
                    with ca:
                        client_sel  = st.selectbox("Client *", list(clients_dict.keys()))
                        fonds_sel   = st.selectbox("Fonds *", FONDS)
                        statut_sel  = st.selectbox("Statut *", STATUTS)
                        owner_input = st.text_input("Commercial (Sales Owner)",
                                                    value=sales_owners[0]
                                                    if sales_owners else "")
                    with cb:
                        target_aum  = st.number_input("AUM Cible (EUR)",
                                                      min_value=0, step=1_000_000)
                        revised_aum = st.number_input("AUM Revise (EUR)",
                                                      min_value=0, step=1_000_000)
                        funded_aum  = st.number_input("AUM Finance (EUR)",
                                                      min_value=0, step=1_000_000)

                    raison_perte, concurrent = "", ""
                    if statut_sel in ("Lost", "Paused"):
                        cc, cd = st.columns(2)
                        with cc:
                            raison_perte = st.selectbox("Raison *", RAISONS_PERTE)
                        with cd:
                            concurrent = st.text_input("Concurrent")

                    next_action = st.date_input(
                        "Prochaine Action",
                        value=date.today() + timedelta(days=14)
                    )
                    sub_d = st.form_submit_button("Enregistrer le Deal",
                                                  use_container_width=True)

                if sub_d:
                    if statut_sel in ("Lost", "Paused") and not raison_perte:
                        st.error("Raison de perte/pause obligatoire.")
                    else:
                        db.add_pipeline_entry(
                            clients_dict[client_sel], fonds_sel, statut_sel,
                            float(target_aum), float(revised_aum), float(funded_aum),
                            raison_perte, concurrent, next_action.isoformat(),
                            owner_input.strip() or "Non assigne"
                        )
                        st.success(f"Deal {fonds_sel} / {client_sel} enregistre.")

        with st.expander("Enregistrer une Activite"):
            clients_dict2 = db.get_client_options()
            if clients_dict2:
                with st.form("form_act", clear_on_submit=True):
                    ce, cf = st.columns(2)
                    with ce:
                        act_client = st.selectbox("Client *",
                                                  list(clients_dict2.keys()))
                        act_type   = st.selectbox("Type", TYPES_INTERACTION)
                    with cf:
                        act_date   = st.date_input("Date", value=date.today())
                        act_notes  = st.text_area("Notes", height=70)
                    sub_a = st.form_submit_button("Enregistrer", use_container_width=True)
                if sub_a:
                    db.add_activity(clients_dict2[act_client],
                                    act_date.isoformat(), act_notes, act_type)
                    st.success(f"Activite enregistree pour {act_client}.")

    with col_import:
        st.markdown("#### Import CSV / Excel — Upsert")
        import_type = st.radio("Table cible", ["Clients", "Pipeline"],
                               horizontal=True)

        if import_type == "Clients":
            st.info("Colonnes : nom_client, type_client, region")
        else:
            st.info("Colonnes : nom_client, fonds, statut, target_aum_initial, "
                    "revised_aum, funded_aum, raison_perte, concurrent_choisi, "
                    "next_action_date, sales_owner")

        uploaded_file = st.file_uploader(
            "Fichier CSV ou Excel (.xlsx)",
            type=["csv","xlsx","xls"]
        )
        if uploaded_file:
            try:
                df_imp = (pd.read_csv(uploaded_file)
                          if uploaded_file.name.endswith(".csv")
                          else pd.read_excel(uploaded_file))
                st.dataframe(df_imp.head(5), use_container_width=True, height=150)
                st.caption(f"{len(df_imp)} ligne(s) detectee(s)")
                if st.button("Lancer l'import (Upsert)", use_container_width=True):
                    with st.spinner("Import en cours..."):
                        fn = (db.upsert_clients_from_df
                              if import_type == "Clients"
                              else db.upsert_pipeline_from_df)
                        ins, upd = fn(df_imp)
                    st.success(f"Import termine : {ins} cree(s), {upd} mis a jour.")
            except Exception as e:
                st.error(f"Erreur de lecture : {e}")

        st.divider()
        st.markdown("#### Dernieres Activites")
        df_act = db.get_activities()
        if not df_act.empty:
            st.dataframe(
                df_act[["nom_client","date","type_interaction","notes"]].head(10),
                use_container_width=True, height=260, hide_index=True,
                column_config={
                    "nom_client":       st.column_config.TextColumn("Client"),
                    "date":             st.column_config.TextColumn("Date"),
                    "type_interaction": st.column_config.TextColumn("Type"),
                    "notes":            st.column_config.TextColumn("Notes"),
                }
            )


# ============================================================================
# ONGLET 2 : PIPELINE MANAGEMENT — MASTER-DETAIL + AUDIT TRAIL
# ============================================================================
with tab_pipeline:
    st.markdown('<div class="section-title">Pipeline Management</div>',
                unsafe_allow_html=True)

    # Filtres
    with st.expander("Filtres", expanded=False):
        fc1, fc2, fc3 = st.columns(3)
        with fc1:
            filt_statuts = st.multiselect("Statuts", STATUTS, default=STATUTS_ACTIFS)
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
        f'<div class="pipeline-hint">Selectionnez une ligne pour ouvrir le panneau '
        f'de modification — <b>{len(df_view)} deal(s)</b> affiches</div>',
        unsafe_allow_html=True
    )

    # Vue lecture seule avec dates converties en str
    df_display = df_view.copy()
    df_display["next_action_date"] = df_display["next_action_date"].apply(
        lambda d: d.isoformat() if isinstance(d, date) else ""
    )
    cols_show = ["id", "nom_client", "type_client", "region", "fonds", "statut",
                 "target_aum_initial", "revised_aum", "funded_aum",
                 "raison_perte", "concurrent_choisi", "next_action_date",
                 "sales_owner"]

    event = st.dataframe(
        df_display[cols_show],
        use_container_width=True,
        height=380,
        hide_index=True,
        on_select="rerun",
        selection_mode="single-row",
        column_config={
            "id":                 st.column_config.NumberColumn("ID", width="small"),
            "nom_client":         st.column_config.TextColumn("Client"),
            "type_client":        st.column_config.TextColumn("Type", width="small"),
            "region":             st.column_config.TextColumn("Region", width="small"),
            "fonds":              st.column_config.TextColumn("Fonds"),
            "statut":             st.column_config.TextColumn("Statut"),
            "target_aum_initial": st.column_config.NumberColumn("AUM Cible",
                                                                 format="EUR %,.0f"),
            "revised_aum":        st.column_config.NumberColumn("AUM Revise",
                                                                 format="EUR %,.0f"),
            "funded_aum":         st.column_config.NumberColumn("AUM Finance",
                                                                 format="EUR %,.0f"),
            "raison_perte":       st.column_config.TextColumn("Raison"),
            "concurrent_choisi":  st.column_config.TextColumn("Concurrent"),
            "next_action_date":   st.column_config.TextColumn("Next Action"),
            "sales_owner":        st.column_config.TextColumn("Commercial"),
        },
        key="pipeline_readonly"
    )

    selected_rows = event.selection.rows if event.selection else []

    if len(selected_rows) > 0:
        sel_idx     = selected_rows[0]
        sel_row     = df_view.iloc[sel_idx]
        pipeline_id = int(sel_row["id"])
        row_data    = db.get_pipeline_row_by_id(pipeline_id)

        if row_data:
            client_name    = str(row_data.get("nom_client", ""))
            current_statut = str(row_data.get("statut", "Prospect"))

            st.markdown(f"""
            <div class="detail-panel">
                <div style="font-size:0.95rem;font-weight:700;color:{C_MARINE};
                            margin-bottom:14px;">
                    Modification du Deal — 
                    <span style="color:{C_CIEL};">{client_name}</span>
                    &nbsp; {statut_badge_html(current_statut)}
                    &nbsp;
                    <span style="font-size:0.75rem;color:#888;font-weight:400;">
                        Deal ID #{pipeline_id}
                    </span>
                </div>
            </div>
            """, unsafe_allow_html=True)

            with st.container():
                with st.form(key=f"edit_deal_{pipeline_id}"):
                    r1c1, r1c2 = st.columns(2)
                    with r1c1:
                        fonds_idx  = (FONDS.index(row_data["fonds"])
                                      if row_data["fonds"] in FONDS else 0)
                        new_fonds  = st.selectbox("Fonds", FONDS, index=fonds_idx)
                    with r1c2:
                        stat_idx   = (STATUTS.index(current_statut)
                                      if current_statut in STATUTS else 0)
                        new_statut = st.selectbox("Statut", STATUTS, index=stat_idx)

                    r2c1, r2c2, r2c3 = st.columns(3)
                    with r2c1:
                        new_target = st.number_input(
                            "AUM Cible (EUR)",
                            value=float(row_data.get("target_aum_initial", 0.0)),
                            min_value=0.0, step=1_000_000.0, format="%.0f"
                        )
                    with r2c2:
                        new_revised = st.number_input(
                            "AUM Revise (EUR)",
                            value=float(row_data.get("revised_aum", 0.0)),
                            min_value=0.0, step=1_000_000.0, format="%.0f"
                        )
                    with r2c3:
                        new_funded = st.number_input(
                            "AUM Finance (EUR)",
                            value=float(row_data.get("funded_aum", 0.0)),
                            min_value=0.0, step=1_000_000.0, format="%.0f"
                        )

                    r3c1, r3c2, r3c3, r3c4 = st.columns(4)
                    with r3c1:
                        raison_opts  = [""] + RAISONS_PERTE
                        cur_raison   = str(row_data.get("raison_perte") or "")
                        raison_idx   = (raison_opts.index(cur_raison)
                                        if cur_raison in raison_opts else 0)
                        lbl_raison   = ("Raison (obligatoire)"
                                        if new_statut in ("Lost","Paused")
                                        else "Raison Perte/Pause")
                        new_raison   = st.selectbox(lbl_raison, raison_opts,
                                                    index=raison_idx)
                    with r3c2:
                        new_concurrent = st.text_input(
                            "Concurrent Choisi",
                            value=str(row_data.get("concurrent_choisi") or "")
                        )
                    with r3c3:
                        nad_val = row_data.get("next_action_date")
                        if not isinstance(nad_val, date):
                            nad_val = date.today() + timedelta(days=14)
                        new_nad = st.date_input("Prochaine Action", value=nad_val)
                    with r3c4:
                        new_sales = st.text_input(
                            "Commercial",
                            value=str(row_data.get("sales_owner") or "Non assigne")
                        )

                    submitted = st.form_submit_button(
                        "Sauvegarder les modifications",
                        use_container_width=False
                    )

                if submitted:
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

            # --- Historique des modifications (Audit Trail) ---
            with st.expander(
                f"Historique des modifications — Deal #{pipeline_id}",
                expanded=False
            ):
                df_audit = db.get_audit_log(pipeline_id)
                if df_audit.empty:
                    st.info("Aucune modification enregistree pour ce deal.")
                else:
                    st.dataframe(
                        df_audit,
                        use_container_width=True,
                        hide_index=True,
                        height=min(250, 45 + len(df_audit) * 38),
                        column_config={
                            "Champ":          st.column_config.TextColumn("Champ"),
                            "Ancienne valeur":st.column_config.TextColumn("Avant"),
                            "Nouvelle valeur":st.column_config.TextColumn("Apres"),
                            "Modifie par":    st.column_config.TextColumn("Operateur"),
                            "Date":           st.column_config.TextColumn("Date/Heure"),
                        }
                    )

    else:
        st.markdown(f"""
        <div style="background:{C_MARINE}05;border:1px dashed {C_MARINE}25;
                    border-radius:8px;padding:24px;text-align:center;margin-top:12px;">
            <div style="color:{C_MARINE};font-weight:600;font-size:0.9rem;">
                Selectionnez un deal dans le tableau pour ouvrir le formulaire
                de modification
            </div>
            <div style="color:#888;font-size:0.8rem;margin-top:4px;">
                Le formulaire et l'historique des modifications apparaitront ici
            </div>
        </div>
        """, unsafe_allow_html=True)

    st.divider()

    # Graphique AUM groupes
    st.markdown("#### Comparaison AUM par Deal")
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    df_viz = df_pipe[
        (df_pipe["target_aum_initial"] > 0) &
        (df_pipe["statut"].isin(["Funded","Soft Commit","Due Diligence","Initial Pitch"]))
    ].copy().head(10)

    if not df_viz.empty:
        fig, ax = plt.subplots(figsize=(12, 4.0))
        fig.patch.set_facecolor("white")
        ax.set_facecolor("white")
        x = np.arange(len(df_viz))
        w = 0.28
        ax.bar(x - w,   df_viz["target_aum_initial"], w,
               label="AUM Cible",   color=C_GRIS,     edgecolor="white")
        ax.bar(x,       df_viz["revised_aum"],  w,
               label="AUM Revise",  color=C_BLEU_MID, edgecolor="white")
        ax.bar(x + w,   df_viz["funded_aum"],   w,
               label="AUM Finance", color=C_CIEL,     edgecolor="white")
        ax.set_xticks(x)
        ax.set_xticklabels(df_viz["nom_client"].str[:18],
                           rotation=26, ha="right", fontsize=8.5, color=C_MARINE)
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda y,_: f"${y/1e6:.0f}M"))
        ax.tick_params(axis="y", colors=C_MARINE, labelsize=8)
        ax.set_title("AUM Cible / Revise / Finance — Deals Actifs",
                     fontsize=10, fontweight="bold", color=C_MARINE, pad=7)
        ax.legend(fontsize=8.5, frameon=False, labelcolor=C_MARINE)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.grid(axis="y", alpha=0.22, color=C_GRIS)
        fig.tight_layout()
        st.pyplot(fig, use_container_width=True)
        plt.close(fig)
    else:
        st.info("Donnees insuffisantes pour le graphique.")

    # Tableau Lost/Paused
    df_lp = df_pipe[df_pipe["statut"].isin(["Lost","Paused"])].copy()
    if not df_lp.empty:
        st.divider()
        st.markdown("#### Deals Perdus / En Pause")
        df_lp_d = df_lp[["nom_client","fonds","statut","target_aum_initial",
                          "raison_perte","concurrent_choisi","sales_owner"]].copy()
        st.dataframe(df_lp_d, use_container_width=True, hide_index=True,
            column_config={
                "nom_client":         st.column_config.TextColumn("Client"),
                "fonds":              st.column_config.TextColumn("Fonds"),
                "statut":             st.column_config.TextColumn("Statut"),
                "target_aum_initial": st.column_config.NumberColumn("AUM Cible",
                                                                     format="EUR %,.0f"),
                "raison_perte":       st.column_config.TextColumn("Raison"),
                "concurrent_choisi":  st.column_config.TextColumn("Concurrent"),
                "sales_owner":        st.column_config.TextColumn("Commercial"),
            })


# ============================================================================
# ONGLET 3 : EXECUTIVE DASHBOARD
# ============================================================================
with tab_dash:
    st.markdown('<div class="section-title">Executive Dashboard</div>',
                unsafe_allow_html=True)
    kpis = db.get_kpis()

    # KPIs
    kpi_cols = st.columns(4, gap="medium")
    for col, (lbl, val, sub) in zip(kpi_cols, [
        ("AUM Finance Total",  fmt_m(kpis["total_funded"]),
         f"{kpis['nb_funded']} deal(s) Funded"),
        ("Pipeline Actif",     fmt_m(kpis["pipeline_actif"]),
         f"{kpis['nb_deals_actifs']} en cours"),
        ("Taux Conversion",    f"{kpis['taux_conversion']}%",
         f"{kpis['nb_funded']} funded / {kpis['nb_lost']} lost"),
        ("Deals Actifs",       str(kpis["nb_deals_actifs"]),
         "Prospect a Soft Commit"),
    ]):
        with col: st.markdown(kpi_card(lbl, val, sub), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    STATUT_COLORS = {
        "Funded": C_CIEL, "Soft Commit": C_BLEU_MID, "Due Diligence": "#004F8C",
        "Initial Pitch": "#3A7EBA", "Prospect": "#8FBEDB",
        "Lost": "#AAAAAA", "Paused": "#C0D8E8", "Redeemed": "#B0CEE8",
    }
    statut_order = [s for s in ["Prospect","Initial Pitch","Due Diligence",
                                 "Soft Commit","Funded","Lost","Paused","Redeemed"]
                    if kpis["statut_repartition"].get(s, 0) > 0]
    if statut_order:
        bcols = st.columns(len(statut_order), gap="small")
        for col, statut in zip(bcols, statut_order):
            c_hex = STATUT_COLORS.get(statut, C_CIEL)
            count = kpis["statut_repartition"][statut]
            with col:
                st.markdown(f"""
                <div style="background:{c_hex}1A;border:1px solid {c_hex}44;
                            border-radius:7px;padding:9px;text-align:center;">
                    <div style="font-size:0.66rem;color:{C_MARINE};font-weight:700;
                                text-transform:uppercase;">{statut}</div>
                    <div style="font-size:1.5rem;font-weight:800;color:{c_hex};">{count}</div>
                </div>""", unsafe_allow_html=True)

    st.divider()

    # Alertes
    df_overdue = db.get_overdue_actions()
    if not df_overdue.empty:
        st.markdown(f"""
        <div style="background:#FFF8E1;border-left:4px solid {C_CIEL};
                    border-radius:0 7px 7px 0;padding:11px 15px;margin-bottom:14px;">
            <div style="font-size:0.84rem;font-weight:700;color:{C_MARINE};
                        margin-bottom:7px;">
                {len(df_overdue)} action(s) en retard — intervention requise
            </div>
        """, unsafe_allow_html=True)
        for _, row in df_overdue.iterrows():
            nad = row.get("next_action_date")
            if isinstance(nad, date):
                days_late = (date.today() - nad).days
                nad_str   = nad.isoformat()
            else:
                days_late = 0
                nad_str   = str(nad or "—")
            owner_str = str(row.get("sales_owner", "")) or ""
            st.markdown(f"""
            <div class="alert-overdue">
                <b>{row['nom_client']}</b> — {row['fonds']}
                <span style="color:{C_CIEL};font-weight:600;"> ({row['statut']})</span>
                — Prevue le <b>{nad_str}</b>
                <span style="color:#B04000;"> ({days_late} j. de retard)</span>
                {f' — Commercial : <b>{owner_str}</b>' if owner_str else ''}
            </div>""", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    # Graphiques
    gcol1, gcol2 = st.columns([1, 1.3], gap="large")

    with gcol1:
        st.markdown("#### AUM Funded par Type Client")
        if kpis["aum_by_type"]:
            fig_p, ax_p = plt.subplots(figsize=(5, 4.0))
            fig_p.patch.set_facecolor("white")
            lbls = list(kpis["aum_by_type"].keys())
            vals = list(kpis["aum_by_type"].values())
            pal  = [C_CIEL, C_MARINE, C_BLEU_MID, C_BLEU_PALE, C_BLEU_DEEP]
            cols = [pal[i % len(pal)] for i in range(len(lbls))]
            _, _, autotexts = ax_p.pie(
                vals, colors=cols, autopct="%1.1f%%", startangle=90,
                pctdistance=0.74,
                wedgeprops={"width":0.54,"edgecolor":"white","linewidth":1.5}
            )
            for at in autotexts:
                at.set_fontsize(8.5); at.set_color("white"); at.set_fontweight("bold")
            ax_p.text(0, 0.07, fmt_m(sum(vals)), ha="center", va="center",
                      fontsize=10, fontweight="bold", color=C_MARINE)
            ax_p.text(0, -0.18, "Total Funded", ha="center", va="center",
                      fontsize=7, color="#666")
            patches = [matplotlib.patches.Patch(
                color=cols[i], label=f"{lbls[i]}: {fmt_m(vals[i])}")
                for i in range(len(lbls))]
            ax_p.legend(handles=patches, loc="lower center",
                        bbox_to_anchor=(0.5,-0.22), ncol=2,
                        fontsize=7, frameon=False, labelcolor=C_MARINE)
            fig_p.tight_layout()
            st.pyplot(fig_p, use_container_width=True)
            plt.close(fig_p)

    with gcol2:
        st.markdown("#### AUM Funded par Fonds")
        if kpis["aum_by_fonds"]:
            fig_b, ax_b = plt.subplots(figsize=(7, 4.0))
            fig_b.patch.set_facecolor("white"); ax_b.set_facecolor("white")
            flbls = list(kpis["aum_by_fonds"].keys())
            fvals = list(kpis["aum_by_fonds"].values())
            yp    = range(len(flbls))
            bars  = ax_b.barh(yp, fvals, color=C_CIEL, edgecolor="white", height=0.52)
            for bar, val in zip(bars, fvals):
                ax_b.text(bar.get_width() + max(fvals)*0.01,
                          bar.get_y() + bar.get_height()/2,
                          fmt_m(val), va="center", ha="left",
                          fontsize=8.5, color=C_MARINE, fontweight="bold")
            ax_b.set_yticks(yp)
            ax_b.set_yticklabels(flbls, fontsize=8.5, color=C_MARINE)
            ax_b.invert_yaxis()
            ax_b.xaxis.set_major_formatter(
                plt.FuncFormatter(lambda x,_: fmt_m(x))
            )
            ax_b.tick_params(axis="x", colors=C_MARINE, labelsize=7.5)
            ax_b.set_xlim(0, max(fvals)*1.18)
            ax_b.spines["top"].set_visible(False)
            ax_b.spines["right"].set_visible(False)
            ax_b.grid(axis="x", alpha=0.22, color=C_GRIS)
            fig_b.tight_layout()
            st.pyplot(fig_b, use_container_width=True)
            plt.close(fig_b)

    st.divider()
    st.markdown("#### Top Deals — AUM Finance")
    df_funded = (db.get_pipeline_with_clients()
                 .query("statut == 'Funded'")
                 .sort_values("funded_aum", ascending=False)
                 .head(10))

    if not df_funded.empty:
        max_f = float(df_funded["funded_aum"].max())
        for i, (_, row) in enumerate(df_funded.iterrows()):
            val  = float(row["funded_aum"])
            pct  = val / max_f * 100 if max_f > 0 else 0
            rank = ["No.1", "No.2", "No.3"][i] if i < 3 else f"No.{i+1}"
            st.markdown(f"""
            <div style="display:flex;align-items:center;gap:11px;margin:6px 0;
                        padding:9px 13px;background:#F8FAFD;border-radius:7px;
                        border:1px solid {C_MARINE}0D;">
                <div style="font-size:0.78rem;font-weight:700;color:{C_BLEU_MID};
                            min-width:32px;">{rank}</div>
                <div style="flex:1;">
                    <div style="font-size:0.85rem;font-weight:700;color:{C_MARINE};">
                        {row['nom_client']}</div>
                    <div style="font-size:0.72rem;color:#777;">
                        {row['fonds']} &middot; {row['type_client']}
                        &middot; {row['region']}</div>
                    <div style="background:{C_GRIS};border-radius:3px;height:5px;
                                margin-top:4px;overflow:hidden;">
                        <div style="background:{C_CIEL};width:{pct:.0f}%;
                                    height:100%;border-radius:3px;"></div>
                    </div>
                </div>
                <div style="text-align:right;min-width:72px;">
                    <div style="font-size:0.92rem;font-weight:800;color:{C_CIEL};">
                        {fmt_m(val)}</div>
                    <div style="font-size:0.66rem;color:#999;">finance</div>
                </div>
            </div>""", unsafe_allow_html=True)
    else:
        st.info("Aucun deal en statut Funded.")


# ============================================================================
# ONGLET 4 : SALES TRACKING
# ============================================================================
with tab_sales:
    st.markdown('<div class="section-title">Sales Tracking — Suivi par Commercial</div>',
                unsafe_allow_html=True)

    df_sales_metrics = db.get_sales_metrics()
    df_next_actions  = db.get_next_actions_by_sales(days_ahead=30)

    if df_sales_metrics.empty:
        st.info("Aucune donnee de pipeline disponible.")
    else:
        # Cartes par commercial
        n_owners = len(df_sales_metrics)
        s_cols   = st.columns(min(n_owners, 3), gap="medium")

        for i, (_, row) in enumerate(df_sales_metrics.iterrows()):
            col_idx = i % min(n_owners, 3)
            with s_cols[col_idx]:
                retard_badge = (
                    f'<span style="color:#B04000;font-weight:700;">'
                    f'{int(row["Actions en retard"])} action(s) en retard</span>'
                    if int(row["Actions en retard"]) > 0
                    else '<span style="color:#2E7D32;">A jour</span>'
                )
                st.markdown(f"""
                <div class="sales-card">
                    <div class="sales-card-name">{row['Commercial']}</div>
                    <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;">
                        <div>
                            <div class="sales-metric">Deals Total</div>
                            <div class="sales-metric-val">{int(row['Nb Deals'])}</div>
                        </div>
                        <div>
                            <div class="sales-metric">Funded</div>
                            <div class="sales-metric-val">{int(row['Funded'])}</div>
                        </div>
                        <div>
                            <div class="sales-metric">AUM Finance</div>
                            <div class="sales-metric-ciel">
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
                            <div style="font-size:0.78rem;">{retard_badge}</div>
                        </div>
                    </div>
                </div>
                """, unsafe_allow_html=True)
                st.markdown("<br>", unsafe_allow_html=True)

        st.divider()

        # Tableau recapitulatif
        st.markdown("#### Tableau de bord commercial")
        df_sales_display = df_sales_metrics.copy()
        for col in ["AUM Finance", "Pipeline Actif"]:
            df_sales_display[col] = df_sales_display[col].apply(fmt_m)

        st.dataframe(
            df_sales_display,
            use_container_width=True,
            hide_index=True,
            height=min(350, 50 + len(df_sales_display) * 40),
            column_config={
                "Commercial":        st.column_config.TextColumn("Commercial"),
                "Nb Deals":          st.column_config.NumberColumn("Total Deals"),
                "Funded":            st.column_config.NumberColumn("Funded"),
                "Actifs":            st.column_config.NumberColumn("Actifs"),
                "Perdus":            st.column_config.NumberColumn("Perdus"),
                "AUM Finance":       st.column_config.TextColumn("AUM Finance"),
                "Pipeline Actif":    st.column_config.TextColumn("Pipeline Actif"),
                "Actions en retard": st.column_config.NumberColumn("Retards"),
            }
        )

        st.divider()

        # Graphique AUM finance par commercial
        st.markdown("#### AUM Finance par Commercial")
        if not df_sales_metrics.empty and df_sales_metrics["AUM Finance"].sum() > 0:
            fig_s, ax_s = plt.subplots(figsize=(10, 3.5))
            fig_s.patch.set_facecolor("white")
            ax_s.set_facecolor("white")

            owners_list = df_sales_metrics["Commercial"].tolist()
            aum_vals    = df_sales_metrics["AUM Finance"].tolist()
            pipe_vals   = df_sales_metrics["Pipeline Actif"].tolist()

            x = np.arange(len(owners_list))
            w = 0.36
            ax_s.bar(x - w/2, aum_vals,  w, label="AUM Finance",
                     color=C_CIEL,    edgecolor="white")
            ax_s.bar(x + w/2, pipe_vals, w, label="Pipeline Actif",
                     color=C_BLEU_MID, edgecolor="white")

            ax_s.set_xticks(x)
            ax_s.set_xticklabels(owners_list, fontsize=9, color=C_MARINE)
            ax_s.yaxis.set_major_formatter(
                plt.FuncFormatter(lambda y,_: fmt_m(y))
            )
            ax_s.tick_params(colors=C_MARINE, labelsize=8)
            ax_s.set_title("AUM par Commercial — Finance vs Pipeline Actif",
                           fontsize=10, fontweight="bold", color=C_MARINE, pad=7)
            ax_s.legend(fontsize=8.5, frameon=False, labelcolor=C_MARINE)
            ax_s.spines["top"].set_visible(False)
            ax_s.spines["right"].set_visible(False)
            ax_s.grid(axis="y", alpha=0.22, color=C_GRIS)
            fig_s.tight_layout()
            st.pyplot(fig_s, use_container_width=True)
            plt.close(fig_s)

        st.divider()

        # Prochaines actions par commercial
        st.markdown("#### Prochaines Actions — 30 jours (par commercial)")
        if df_next_actions.empty:
            st.info("Aucune action planifiee dans les 30 prochains jours.")
        else:
            owners_na = sorted(df_next_actions["sales_owner"].unique().tolist())
            filter_owner = st.selectbox(
                "Filtrer par commercial", ["Tous"] + owners_na
            )

            df_na_view = (df_next_actions
                          if filter_owner == "Tous"
                          else df_next_actions[
                              df_next_actions["sales_owner"] == filter_owner
                          ])

            today = date.today()
            for _, row in df_na_view.iterrows():
                nad = row.get("next_action_date")
                if isinstance(nad, date):
                    delta = (nad - today).days
                    if delta < 0:
                        timing = f"RETARD ({abs(delta)} j.)"
                        dot_color = "#B04000"
                    elif delta == 0:
                        timing = "Aujourd'hui"
                        dot_color = C_BLEU_MID
                    elif delta <= 7:
                        timing = f"Dans {delta} j."
                        dot_color = C_CIEL
                    else:
                        timing = f"Dans {delta} j."
                        dot_color = C_MARINE
                    nad_str = nad.isoformat()
                else:
                    timing    = "—"
                    dot_color = C_GRIS
                    nad_str   = "—"

                revised = float(row.get("revised_aum", 0) or 0)
                st.markdown(f"""
                <div style="display:flex;align-items:center;gap:12px;padding:8px 12px;
                            margin:4px 0;background:#F8FAFD;border-radius:6px;
                            border-left:3px solid {dot_color};">
                    <div style="min-width:110px;font-size:0.76rem;
                                color:{dot_color};font-weight:700;">{timing}</div>
                    <div style="flex:1;">
                        <span style="font-weight:600;color:{C_MARINE};font-size:0.84rem;">
                            {row['nom_client']}</span>
                        <span style="color:#888;font-size:0.76rem;">
                            &nbsp; {row['fonds']} &middot; {row['statut']}</span>
                    </div>
                    <div style="font-size:0.8rem;color:{C_BLEU_MID};font-weight:600;
                                min-width:70px;text-align:right;">{fmt_m(revised)}</div>
                    <div style="font-size:0.76rem;color:#888;min-width:90px;
                                text-align:right;">{row['sales_owner']}</div>
                </div>""", unsafe_allow_html=True)


# ============================================================================
# ONGLET 5 : PERFORMANCE ET NAV
# ============================================================================
with tab_perf:
    st.markdown('<div class="section-title">Performance et NAV</div>',
                unsafe_allow_html=True)

    col_up, col_info = st.columns([1.2, 1], gap="large")

    with col_info:
        st.markdown(f"""
        <div style="background:{C_MARINE}07;border:1px solid {C_MARINE}15;
                    border-radius:9px;padding:16px 18px;">
            <div style="font-weight:700;color:{C_MARINE};font-size:0.92rem;
                        margin-bottom:10px;">Format attendu</div>
            <div style="font-size:0.81rem;color:#444;line-height:1.75;">
                <b>Colonnes obligatoires</b><br>
                &bull; <code>Date</code> — YYYY-MM-DD ou DD/MM/YYYY<br>
                &bull; <code>Fonds</code> — Nom du fonds<br>
                &bull; <code>NAV</code> — Valeur liquidative numerique<br><br>
                <b>Calculs produits</b><br>
                &bull; Courbe Base 100 normalisee au premier point<br>
                &bull; Performance 1 Mois glissant<br>
                &bull; Performance YTD depuis le 1er janvier<br>
                &bull; Performance sur la periode selectionnee<br><br>
                <b>Export PDF</b><br>
                Une fois les donnees chargees, cliquez sur<br>
                <i>Generer le rapport PDF</i> dans la barre laterale
                pour inclure cette page dans l'export.
            </div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("#### Donnees de demonstration")
        if st.button("Generer un fichier NAV de demonstration",
                     use_container_width=True):
            demo_dates = pd.date_range(
                start=f"{date.today().year - 1}-01-01",
                end=date.today(), freq="B"
            )
            rng  = np.random.default_rng(42)
            navs = {"Global Value": 100.0, "Resilient Equity": 100.0,
                    "Income Builder": 100.0, "Private Debt": 100.0,
                    "International Fund": 100.0}
            rows = []
            for d in demo_dates:
                for fonds, nav in navs.items():
                    nav *= (1 + rng.normal(0.0003, 0.006))
                    navs[fonds] = nav
                    rows.append({"Date": d.date().isoformat(),
                                 "Fonds": fonds, "NAV": round(nav, 4)})
            df_demo = pd.DataFrame(rows)
            buf_demo = io.BytesIO()
            df_demo.to_excel(buf_demo, index=False)
            buf_demo.seek(0)
            st.download_button(
                "Telecharger nav_demo.xlsx",
                data=buf_demo, file_name="nav_demo.xlsx",
                mime="application/vnd.ms-excel",
                use_container_width=True
            )

    with col_up:
        nav_file = st.file_uploader(
            "Charger l'historique NAV (Excel ou CSV)",
            type=["xlsx","xls","csv"]
        )

    if nav_file is not None:
        try:
            df_nav = (pd.read_csv(nav_file)
                      if nav_file.name.endswith(".csv")
                      else pd.read_excel(nav_file))

            df_nav.columns = [c.strip() for c in df_nav.columns]

            missing = [c for c in ["Date","Fonds","NAV"] if c not in df_nav.columns]
            if missing:
                st.error(f"Colonnes manquantes : {missing}. "
                         f"Colonnes trouvees : {list(df_nav.columns)}")
                st.stop()

            # --- Nettoyage securise des types (Mission 4) ---
            df_nav["Date"] = pd.to_datetime(
                df_nav["Date"], format="mixed", dayfirst=True, errors="coerce"
            )
            # Supprimer les NaT immediatement
            df_nav = df_nav.dropna(subset=["Date"])

            df_nav["NAV"]   = pd.to_numeric(df_nav["NAV"], errors="coerce")
            df_nav          = df_nav.dropna(subset=["NAV"])
            df_nav["Fonds"] = df_nav["Fonds"].astype(str).str.strip()
            df_nav          = df_nav.sort_values("Date").reset_index(drop=True)

            if df_nav.empty:
                st.error("Aucune donnee valide apres nettoyage. "
                         "Verifiez le format de vos colonnes Date et NAV.")
                st.stop()

            fonds_list = sorted(df_nav["Fonds"].unique().tolist())
            date_min   = df_nav["Date"].min()
            date_max   = df_nav["Date"].max()

            st.markdown(f"""
            <div style="background:{C_CIEL}16;border-left:3px solid {C_CIEL};
                        border-radius:0 7px 7px 0;padding:9px 15px;margin:10px 0;">
                {len(df_nav):,} points charges &mdash; {len(fonds_list)} fonds &mdash;
                Periode : {date_min.strftime('%d/%m/%Y')} au {date_max.strftime('%d/%m/%Y')}
            </div>
            """, unsafe_allow_html=True)

            # Filtres
            st.markdown("---")
            ff1, ff2, ff3 = st.columns([2, 1, 1])
            with ff1:
                fonds_sel_nav = st.multiselect(
                    "Fonds a afficher", fonds_list,
                    default=fonds_list[:min(5, len(fonds_list))]
                )
            with ff2:
                d_debut = st.date_input("Depuis", value=date_min.date())
            with ff3:
                d_fin   = st.date_input("Jusqu'au", value=date_max.date())

            if not fonds_sel_nav:
                st.warning("Selectionnez au moins un fonds.")
                st.stop()

            mask = (
                df_nav["Fonds"].isin(fonds_sel_nav) &
                (df_nav["Date"].dt.date >= d_debut) &
                (df_nav["Date"].dt.date <= d_fin)
            )
            df_filtered_nav = df_nav[mask].copy()

            if df_filtered_nav.empty:
                st.warning("Aucune donnee pour la periode selectionnee.")
                st.stop()

            # Pivot et Base 100
            pivot = (df_filtered_nav
                     .pivot_table(index="Date", columns="Fonds",
                                  values="NAV", aggfunc="last")
                     .sort_index()
                     .ffill())
            pivot = pivot[[f for f in fonds_sel_nav if f in pivot.columns]]

            base100 = pivot.copy() * np.nan
            for fonds in pivot.columns:
                series = pivot[fonds].dropna()
                if series.empty:
                    continue
                first_val = float(series.iloc[0])
                # Protection division par zero (Mission 4)
                if first_val != 0 and not np.isnan(first_val):
                    base100[fonds] = pivot[fonds] / first_val * 100
                # Si first_val == 0, la colonne reste NaN (non affichee)

            # Stocker en session_state pour le PDF
            st.session_state["nav_base100"] = base100
            st.session_state["nav_fonds"]   = fonds_sel_nav

            # --- Graphique Base 100 ---
            st.markdown("#### Evolution NAV — Base 100")

            fig_nav, ax_nav = plt.subplots(figsize=(12, 5))
            fig_nav.patch.set_facecolor("white")
            ax_nav.set_facecolor("white")

            plotted = 0
            for i, fonds in enumerate(pivot.columns):
                series = base100[fonds].dropna()
                if series.empty:
                    continue
                color    = NAV_PALETTE[i % len(NAV_PALETTE)]
                line_sty = "-" if i % 2 == 0 else "--"

                # Protection single-point (Mission 4)
                if len(series) >= 2:
                    ax_nav.plot(series.index, series.values,
                                label=fonds, color=color,
                                linewidth=1.8, linestyle=line_sty, alpha=0.92)
                else:
                    # Un seul point : affichage en scatter uniquement
                    ax_nav.scatter(series.index, series.values,
                                   color=color, s=60, label=fonds, zorder=5)

                if not series.empty:
                    ax_nav.scatter([series.index[-1]], [series.values[-1]],
                                   color=color, s=35, zorder=6)
                plotted += 1

            if plotted == 0:
                st.warning("Aucune donnee Base 100 calculable.")
                st.stop()

            ax_nav.axhline(100, color=C_GRIS, linewidth=0.9, linestyle=":")
            ax_nav.set_ylabel("NAV (Base 100)", fontsize=8.5, color=C_MARINE)
            ax_nav.tick_params(colors=C_MARINE, labelsize=8)
            ax_nav.spines["top"].set_visible(False)
            ax_nav.spines["right"].set_visible(False)
            ax_nav.spines["left"].set_color(C_GRIS)
            ax_nav.spines["bottom"].set_color(C_GRIS)
            ax_nav.grid(axis="y", alpha=0.22, color=C_GRIS, linewidth=0.6)
            ax_nav.grid(axis="x", alpha=0.08, color=C_GRIS, linewidth=0.4)
            ax_nav.legend(fontsize=8.5, frameon=True, framealpha=0.92,
                          edgecolor=C_GRIS, labelcolor=C_MARINE, loc="upper left")
            ax_nav.set_title(
                f"Performance NAV — Base 100 — "
                f"{d_debut.strftime('%d/%m/%Y')} au {d_fin.strftime('%d/%m/%Y')}",
                fontsize=10.5, fontweight="bold", color=C_MARINE, pad=9
            )
            plt.xticks(rotation=18, ha="right", fontsize=7.5)
            fig_nav.tight_layout()
            st.pyplot(fig_nav, use_container_width=True)
            plt.close(fig_nav)

            # --- Calcul des performances ---
            today_ts  = pd.Timestamp(date.today())
            one_m_ago = today_ts - pd.DateOffset(months=1)
            jan_1     = pd.Timestamp(f"{date.today().year}-01-01")

            perf_rows = []
            for fonds in pivot.columns:
                series = pivot[fonds].dropna()
                if series.empty:
                    continue

                nav_last  = float(series.iloc[-1])
                nav_first = float(series.iloc[0])

                s_1m  = series[series.index >= one_m_ago]
                perf_1m = (
                    (nav_last / float(s_1m.iloc[0]) - 1) * 100
                    if len(s_1m) > 0 and float(s_1m.iloc[0]) != 0
                    else float("nan")
                )

                s_ytd  = series[series.index >= jan_1]
                perf_ytd = (
                    (nav_last / float(s_ytd.iloc[0]) - 1) * 100
                    if len(s_ytd) > 0 and float(s_ytd.iloc[0]) != 0
                    else float("nan")
                )

                perf_period = (
                    (nav_last / nav_first - 1) * 100
                    if nav_first != 0 else float("nan")
                )

                b100_series = base100[fonds].dropna()
                nav_b100    = (float(b100_series.iloc[-1])
                               if not b100_series.empty else float("nan"))

                perf_rows.append({
                    "Fonds":             fonds,
                    "NAV Derniere":      round(nav_last, 4),
                    "Base 100 Actuel":   round(nav_b100, 2)
                                         if not np.isnan(nav_b100) else None,
                    "Perf 1M (%)":       round(perf_1m, 2)
                                         if not np.isnan(perf_1m) else None,
                    "Perf YTD (%)":      round(perf_ytd, 2)
                                         if not np.isnan(perf_ytd) else None,
                    "Perf Periode (%)":  round(perf_period, 2)
                                         if not np.isnan(perf_period) else None,
                })

            if perf_rows:
                df_perf_table = pd.DataFrame(perf_rows)
                # Stocker pour le PDF
                st.session_state["perf_data"] = df_perf_table

                # --- Tableau HTML des performances ---
                st.markdown("#### Tableau des Performances")

                def _fmt_perf(val):
                    if val is None or (isinstance(val, float) and np.isnan(val)):
                        return '<span style="color:#999;">n.d.</span>'
                    color = C_CIEL if val >= 0 else "#8B2020"
                    sign  = "+" if val > 0 else ""
                    return (f'<span style="color:{color};font-weight:700;">'
                            f'{sign}{val:.2f}%</span>')

                tbl_html = f"""
                <table style="width:100%;border-collapse:collapse;font-size:0.84rem;">
                    <thead>
                        <tr style="background:{C_MARINE};color:{C_BLANC};">
                            <th style="padding:9px 13px;text-align:left;">Fonds</th>
                            <th style="padding:9px 13px;text-align:right;">
                                NAV Derniere</th>
                            <th style="padding:9px 13px;text-align:right;">
                                Base 100</th>
                            <th style="padding:9px 13px;text-align:right;">
                                Perf 1M</th>
                            <th style="padding:9px 13px;text-align:right;">
                                Perf YTD</th>
                            <th style="padding:9px 13px;text-align:right;">
                                Perf Periode</th>
                        </tr>
                    </thead>
                    <tbody>
                """
                for i, r in enumerate(perf_rows):
                    bg = "#F8FAFD" if i % 2 == 0 else C_BLANC
                    tbl_html += f"""
                        <tr style="background:{bg};border-bottom:1px solid {C_GRIS};">
                            <td style="padding:8px 13px;font-weight:600;
                                       color:{C_MARINE};">{r['Fonds']}</td>
                            <td style="padding:8px 13px;text-align:right;
                                       color:{C_MARINE};">
                                {r['NAV Derniere']:.4f}</td>
                            <td style="padding:8px 13px;text-align:right;
                                       color:{C_MARINE};">
                                {f"{r['Base 100 Actuel']:.2f}" if r['Base 100 Actuel'] else 'n.d.'}</td>
                            <td style="padding:8px 13px;text-align:right;">
                                {_fmt_perf(r['Perf 1M (%)'])}</td>
                            <td style="padding:8px 13px;text-align:right;">
                                {_fmt_perf(r['Perf YTD (%)'])}</td>
                            <td style="padding:8px 13px;text-align:right;">
                                {_fmt_perf(r['Perf Periode (%)'])}</td>
                        </tr>"""
                tbl_html += "</tbody></table>"
                st.markdown(tbl_html, unsafe_allow_html=True)
                st.markdown("<br>", unsafe_allow_html=True)

                csv_perf = df_perf_table.to_csv(index=False).encode("utf-8")
                st.download_button(
                    "Exporter le tableau en CSV",
                    data=csv_perf,
                    file_name=f"performances_{date.today().isoformat()}.csv",
                    mime="text/csv"
                )

                # --- Graphique YTD ---
                ytd_data = [r for r in perf_rows if r["Perf YTD (%)"] is not None]
                if ytd_data:
                    ytd_data.sort(key=lambda r: r["Perf YTD (%)"], reverse=True)
                    st.markdown("#### Comparaison Performances YTD")

                    fig_ytd, ax_ytd = plt.subplots(figsize=(10, 3.4))
                    fig_ytd.patch.set_facecolor("white")
                    ax_ytd.set_facecolor("white")

                    fonds_ytd = [r["Fonds"] for r in ytd_data]
                    vals_ytd  = [r["Perf YTD (%)"] for r in ytd_data]
                    bar_cols  = [C_CIEL if v >= 0 else C_GRIS for v in vals_ytd]

                    bars_ytd = ax_ytd.bar(
                        range(len(fonds_ytd)), vals_ytd,
                        color=bar_cols, edgecolor="white", linewidth=0.5, width=0.6
                    )
                    ax_ytd.axhline(0, color=C_MARINE, linewidth=0.7)
                    for bar, val in zip(bars_ytd, vals_ytd):
                        sign = "+" if val > 0 else ""
                        ypos = val + 0.04 if val >= 0 else val - 0.22
                        ax_ytd.text(
                            bar.get_x() + bar.get_width()/2, ypos,
                            f"{sign}{val:.2f}%",
                            ha="center",
                            va="bottom" if val >= 0 else "top",
                            fontsize=8, color=C_MARINE, fontweight="bold"
                        )
                    ax_ytd.set_xticks(range(len(fonds_ytd)))
                    ax_ytd.set_xticklabels(fonds_ytd, rotation=16, ha="right",
                                           fontsize=8.5, color=C_MARINE)
                    ax_ytd.set_ylabel("YTD (%)", fontsize=8.5, color=C_MARINE)
                    ax_ytd.tick_params(colors=C_MARINE, labelsize=8)
                    ax_ytd.spines["top"].set_visible(False)
                    ax_ytd.spines["right"].set_visible(False)
                    ax_ytd.grid(axis="y", alpha=0.2, color=C_GRIS)
                    ax_ytd.set_title("Performance YTD par Fonds (%)",
                                     fontsize=10, fontweight="bold",
                                     color=C_MARINE, pad=7)
                    fig_ytd.tight_layout()
                    st.pyplot(fig_ytd, use_container_width=True)
                    plt.close(fig_ytd)

        except Exception as e:
            st.error(f"Erreur lors du traitement du fichier NAV : {e}")
            import traceback
            with st.expander("Details de l'erreur"):
                st.code(traceback.format_exc())

    else:
        st.markdown(f"""
        <div style="background:{C_MARINE}05;border:2px dashed {C_MARINE}22;
                    border-radius:12px;padding:46px;text-align:center;margin-top:14px;">
            <div style="font-size:0.98rem;font-weight:700;color:{C_MARINE};
                        margin-bottom:5px;">Module Performance et NAV</div>
            <div style="color:#777;font-size:0.83rem;max-width:400px;
                        margin:0 auto;line-height:1.65;">
                Chargez un fichier Excel ou CSV avec les colonnes
                <code>Date</code>, <code>Fonds</code>, <code>NAV</code>
                pour generer les courbes Base 100 et le tableau de performances.<br><br>
                Utilisez le bouton <b>Generer un fichier NAV de demonstration</b>
                pour un exemple immediatement testable.
            </div>
        </div>
        """, unsafe_allow_html=True)
