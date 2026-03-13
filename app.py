# =============================================================================
# app.py — CRM & Reporting Tool — FEA/Amundi Edition Strategique
# Charte Amundi : Marine #002D54 | Ciel #00A8E1 | Zero emoji
# Missions :
#   - Filtrage fund-by-fund via Sidebar Perimetre
#   - Dashboard avec graphique AUM par Region
#   - Module Performance robuste (matplotlib interactif via st.pyplot)
#   - Export PDF filtre sur perimetre, avec page Performance optionnelle
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
# CONSTANTES
# ---------------------------------------------------------------------------
C_MARINE    = "#002D54"
C_CIEL      = "#00A8E1"
C_GRIS      = "#E0E0E0"
C_BLANC     = "#FFFFFF"
C_BLEU_MID  = "#1A6B9A"
C_BLEU_PALE = "#7BC8E8"
C_BLEU_DEEP = "#003F7A"

TYPES_CLIENT      = ["IFA", "Wholesale", "Instit", "Family Office"]
REGIONS           = db.REGIONS_REFERENTIEL
FONDS             = db.FONDS_REFERENTIEL
STATUTS           = ["Prospect", "Initial Pitch", "Due Diligence",
                     "Soft Commit", "Funded", "Paused", "Lost", "Redeemed"]
STATUTS_ACTIFS    = ["Prospect", "Initial Pitch", "Due Diligence", "Soft Commit"]
RAISONS_PERTE     = ["Pricing", "Track Record", "Macro", "Competitor", "Autre"]
TYPES_INTERACTION = ["Call", "Meeting", "Email", "Roadshow", "Conference", "Autre"]

NAV_PALETTE = [
    C_CIEL, C_MARINE, C_BLEU_MID, C_BLEU_PALE, C_BLEU_DEEP,
    "#5BA3C9", "#2C8FBF", "#004F8C", "#A8D8EE", "#003060",
]


# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="CRM — Asset Management",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ---------------------------------------------------------------------------
# CSS — CHARTE STRICTE, ZERO EMOJI
# ---------------------------------------------------------------------------
st.markdown(f"""
<style>
  .stApp, .main .block-container {{
      background-color:{C_BLANC}; color:{C_MARINE};
      font-family:'Segoe UI', Arial, sans-serif;
  }}
  [data-testid="stSidebar"] {{ background-color:{C_MARINE}; }}
  [data-testid="stSidebar"] * {{ color:{C_BLANC} !important; }}
  [data-testid="stSidebar"] h1,[data-testid="stSidebar"] h2,
  [data-testid="stSidebar"] h3 {{ color:{C_CIEL} !important; }}
  [data-testid="stSidebar"] .stMultiSelect label {{ color:{C_BLANC} !important; }}

  .crm-header {{
      background:linear-gradient(135deg,{C_MARINE} 0%,#003D6B 100%);
      padding:18px 24px; border-radius:8px; margin-bottom:16px;
      border-left:5px solid {C_CIEL};
  }}
  .crm-header h1 {{
      color:{C_BLANC} !important; margin:0;
      font-size:1.55rem; font-weight:700; letter-spacing:-0.2px;
  }}
  .crm-header p {{ color:{C_CIEL}; margin:3px 0 0 0; font-size:0.82rem; }}

  .kpi-card {{
      background:linear-gradient(135deg,{C_MARINE} 0%,#004070 100%);
      padding:15px 11px; border-radius:8px; text-align:center;
      border:1px solid {C_CIEL}33;
  }}
  .kpi-label {{
      font-size:0.68rem; color:{C_CIEL}; text-transform:uppercase;
      letter-spacing:0.8px; margin-bottom:6px; font-weight:600;
  }}
  .kpi-value {{ font-size:1.5rem; font-weight:800; color:{C_BLANC}; }}
  .kpi-sub   {{ font-size:0.66rem; color:{C_GRIS}; margin-top:3px; }}

  .section-title {{
      font-size:0.98rem; font-weight:700; color:{C_MARINE};
      border-bottom:3px solid {C_CIEL}; padding-bottom:5px;
      margin:15px 0 10px 0;
  }}
  .perimetre-badge {{
      display:inline-block; background:{C_CIEL}18; border:1px solid {C_CIEL}44;
      border-radius:4px; padding:3px 9px; font-size:0.74rem;
      color:{C_MARINE}; font-weight:600; margin:2px;
  }}
  .detail-panel {{
      background:linear-gradient(135deg,#F0F6FC 0%,#E8F2FA 100%);
      border:1.5px solid {C_CIEL}55; border-radius:10px;
      padding:18px 20px 14px 20px; margin-top:14px;
      box-shadow:0 3px 14px {C_MARINE}0C;
  }}
  .pipeline-hint {{
      background:{C_CIEL}0E; border-left:3px solid {C_CIEL};
      border-radius:0 5px 5px 0; padding:7px 12px;
      font-size:0.79rem; color:{C_MARINE}; margin-bottom:9px;
  }}
  .alert-overdue {{
      background:#FFF8E1; border-left:4px solid {C_CIEL};
      border-radius:0 5px 5px 0; padding:7px 12px; margin:4px 0;
      font-size:0.79rem; color:{C_MARINE};
  }}
  .sales-card {{
      background:{C_BLANC}; border:1px solid {C_MARINE}15;
      border-radius:8px; padding:15px; border-top:3px solid {C_CIEL};
  }}
  .sales-card-name {{
      font-size:0.92rem; font-weight:700; color:{C_MARINE};
      margin-bottom:9px; padding-bottom:6px; border-bottom:1px solid {C_GRIS};
  }}
  .sales-metric     {{ font-size:0.72rem; color:#666; margin-bottom:2px; }}
  .sales-metric-val {{ font-size:1.0rem; font-weight:700; color:{C_MARINE}; }}
  .sales-metric-ciel{{ font-size:1.0rem; font-weight:700; color:{C_CIEL}; }}
  .stTabs [data-baseweb="tab-list"] {{
      background:{C_BLANC}; border-bottom:2px solid {C_MARINE}15; gap:2px;
  }}
  .stTabs [data-baseweb="tab"] {{
      color:{C_MARINE}; font-weight:600; font-size:0.82rem;
      padding:8px 16px; border-radius:5px 5px 0 0;
      background:{C_MARINE}07;
  }}
  .stTabs [aria-selected="true"] {{
      background:{C_MARINE} !important; color:{C_BLANC} !important;
  }}
  .stButton > button {{
      background:{C_MARINE}; color:{C_BLANC}; border:none;
      border-radius:5px; font-weight:600; padding:7px 16px;
      font-size:0.83rem; transition:all 0.16s;
  }}
  .stButton > button:hover {{
      background:{C_CIEL}; color:{C_BLANC};
  }}
  .stSelectbox label,.stTextInput label,.stNumberInput label,
  .stDateInput label,.stTextArea label,.stRadio label {{
      color:{C_MARINE} !important; font-weight:600; font-size:0.79rem;
  }}
  h1,h2,h3,h4 {{ color:{C_MARINE} !important; }}
  hr {{ border-color:{C_MARINE}12; }}
  code {{ background:{C_MARINE}09; color:{C_MARINE}; border-radius:3px; }}
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


def statut_badge(statut: str) -> str:
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
    return (f'<span style="padding:2px 9px;border-radius:9px;'
            f'font-size:0.74rem;font-weight:600;'
            f'background:{bg};color:{fg};">{statut}</span>')


# ---------------------------------------------------------------------------
# SIDEBAR — PERIMETRE D'EXPORT (MISSION 1)
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown(f"""
    <div style="text-align:center;padding:6px 0 16px 0;">
        <div style="font-size:0.95rem;font-weight:800;color:{C_CIEL};
                    letter-spacing:0.4px;">
            CRM Asset Management
        </div>
        <div style="font-size:0.68rem;color:{C_GRIS};margin-top:3px;">
            {date.today().strftime("%d %B %Y")}
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.divider()

    # Snapshot global (sans filtre)
    kpis_global = db.get_kpis()
    st.markdown(
        f'<div style="color:{C_CIEL};font-size:0.7rem;font-weight:700;'
        f'text-transform:uppercase;letter-spacing:1px;margin-bottom:6px;">'
        f'Apercu Global</div>',
        unsafe_allow_html=True
    )
    for lbl, val in [
        ("AUM Finance Total", fmt_m(kpis_global["total_funded"])),
        ("Pipeline Actif",    fmt_m(kpis_global["pipeline_actif"])),
    ]:
        st.markdown(f"""
        <div style="background:{C_BLANC}15;padding:9px;border-radius:6px;
                    margin-bottom:5px;">
            <div style="font-size:0.66rem;color:{C_GRIS};">{lbl}</div>
            <div style="font-size:1.15rem;font-weight:800;color:{C_CIEL};">{val}</div>
        </div>""", unsafe_allow_html=True)

    st.divider()

    # --- PERIMETRE D'EXPORT ---
    st.markdown(
        f'<div style="color:{C_CIEL};font-size:0.7rem;font-weight:700;'
        f'text-transform:uppercase;letter-spacing:1px;margin-bottom:6px;">'
        f'Export PDF</div>',
        unsafe_allow_html=True
    )

    fonds_perimetre = st.multiselect(
        "Perimetre de l'export",
        options=FONDS,
        default=FONDS,
        help=(
            "Selectionnez les fonds a inclure dans le rapport PDF. "
            "Tous les KPIs, graphiques et tableaux seront recalcules "
            "sur ce perimetre uniquement."
        ),
        key="fonds_perimetre_select"
    )

    # Perimetre effectif : None si tout selectionne (pas de filtrage SQL)
    _filtre_effectif = fonds_perimetre if (
        fonds_perimetre and len(fonds_perimetre) < len(FONDS)
    ) else None

    mode_comex = st.toggle(
        "Mode Comex — Anonymisation",
        value=False,
        help="Remplace les noms clients par Type-Region dans le PDF."
    )

    if fonds_perimetre and len(fonds_perimetre) < len(FONDS):
        badges = " ".join(
            f'<span class="perimetre-badge">{f}</span>'
            for f in fonds_perimetre
        )
        st.markdown(f"<div>{badges}</div>", unsafe_allow_html=True)
    elif not fonds_perimetre:
        st.warning("Selectionnez au moins un fonds.")

    # Indicateur donnees NAV disponibles
    perf_data_pdf   = st.session_state.get("perf_data",   None)
    nav_base100_pdf = st.session_state.get("nav_base100", None)
    if perf_data_pdf is not None:
        nb_fonds_nav = len(perf_data_pdf)
        st.caption(f"Donnees NAV disponibles : {nb_fonds_nav} fonds.")

    if not fonds_perimetre:
        st.button("Generer le rapport PDF", disabled=True,
                  use_container_width=True)
    elif st.button("Generer le rapport PDF", use_container_width=True):
        with st.spinner("Calcul du perimetre et generation en cours..."):
            try:
                # Recalcul sur le perimetre exact (MISSION 1)
                pipeline_pdf    = db.get_pipeline_with_clients(
                                      fonds_filter=_filtre_effectif)
                kpis_pdf        = db.get_kpis(fonds_filter=_filtre_effectif)
                aum_region_pdf  = db.get_aum_by_region(
                                      fonds_filter=_filtre_effectif)

                # Filtrage NAV sur le perimetre (MISSION 4)
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
                    data=pdf_bytes, file_name=fname,
                    mime="application/pdf",
                    use_container_width=True
                )
                st.success("Rapport genere.")
            except Exception as e:
                st.error(f"Erreur : {e}")

    st.divider()
    st.caption("Version 4.0 — FEA/Amundi Edition Strategique")


# ---------------------------------------------------------------------------
# EN-TETE
# ---------------------------------------------------------------------------
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

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
                with st.form("form_deal", clear_on_submit=True):
                    ca, cb = st.columns(2)
                    with ca:
                        client_sel  = st.selectbox("Client", list(clients_dict.keys()))
                        fonds_sel   = st.selectbox("Fonds", FONDS)
                        statut_sel  = st.selectbox("Statut", STATUTS)
                        owner_input = st.text_input(
                            "Commercial",
                            value=sales_owners[0] if sales_owners else "")
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
                            raison_perte = st.selectbox("Raison", RAISONS_PERTE)
                        with cd:
                            concurrent = st.text_input("Concurrent")
                    next_action = st.date_input(
                        "Prochaine Action",
                        value=date.today() + timedelta(days=14))
                    sub_d = st.form_submit_button("Enregistrer le Deal",
                                                  use_container_width=True)
                if sub_d:
                    if statut_sel in ("Lost", "Paused") and not raison_perte:
                        st.error("Raison obligatoire pour ce statut.")
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
                        act_client = st.selectbox("Client", list(clients_dict2.keys()))
                        act_type   = st.selectbox("Type", TYPES_INTERACTION)
                    with cf:
                        act_date   = st.date_input("Date", value=date.today())
                        act_notes  = st.text_area("Notes", height=68)
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
            "Fichier CSV ou Excel", type=["csv","xlsx","xls"])
        if uploaded_file:
            try:
                df_imp = (pd.read_csv(uploaded_file)
                          if uploaded_file.name.endswith(".csv")
                          else pd.read_excel(uploaded_file))
                st.dataframe(df_imp.head(5), use_container_width=True, height=145)
                st.caption(f"{len(df_imp)} ligne(s)")
                if st.button("Lancer l'import", use_container_width=True):
                    fn = (db.upsert_clients_from_df
                          if import_type == "Clients"
                          else db.upsert_pipeline_from_df)
                    ins, upd = fn(df_imp)
                    st.success(f"Import : {ins} cree(s), {upd} mis a jour.")
            except Exception as e:
                st.error(f"Erreur : {e}")

        st.divider()
        st.markdown("#### Dernieres Activites")
        df_act = db.get_activities()
        if not df_act.empty:
            st.dataframe(
                df_act[["nom_client","date","type_interaction","notes"]].head(10),
                use_container_width=True, height=250, hide_index=True,
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
        f'<div class="pipeline-hint">Selectionnez une ligne pour modifier '
        f'— <b>{len(df_view)} deal(s)</b> affiches</div>',
        unsafe_allow_html=True
    )

    df_display = df_view.copy()
    df_display["next_action_date"] = df_display["next_action_date"].apply(
        lambda d: d.isoformat() if isinstance(d, date) else ""
    )
    cols_show = ["id","nom_client","type_client","region","fonds","statut",
                 "target_aum_initial","revised_aum","funded_aum",
                 "raison_perte","concurrent_choisi","next_action_date","sales_owner"]

    event = st.dataframe(
        df_display[cols_show],
        use_container_width=True, height=370, hide_index=True,
        on_select="rerun", selection_mode="single-row",
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
                <div style="font-size:0.93rem;font-weight:700;color:{C_MARINE};
                            margin-bottom:13px;">
                    Modification — <span style="color:{C_CIEL};">{client_name}</span>
                    &nbsp; {statut_badge(current_statut)}
                    &nbsp;<span style="font-size:0.72rem;color:#888;font-weight:400;">
                        ID #{pipeline_id}</span>
                </div>
            </div>
            """, unsafe_allow_html=True)

            with st.container():
                with st.form(key=f"edit_{pipeline_id}"):
                    r1c1, r1c2 = st.columns(2)
                    with r1c1:
                        fi = FONDS.index(row_data["fonds"]) \
                             if row_data["fonds"] in FONDS else 0
                        new_fonds  = st.selectbox("Fonds", FONDS, index=fi)
                    with r1c2:
                        si = STATUTS.index(current_statut) \
                             if current_statut in STATUTS else 0
                        new_statut = st.selectbox("Statut", STATUTS, index=si)

                    r2c1, r2c2, r2c3 = st.columns(3)
                    with r2c1:
                        new_target = st.number_input(
                            "AUM Cible (EUR)",
                            value=float(row_data.get("target_aum_initial", 0.0)),
                            min_value=0.0, step=1_000_000.0, format="%.0f")
                    with r2c2:
                        new_revised = st.number_input(
                            "AUM Revise (EUR)",
                            value=float(row_data.get("revised_aum", 0.0)),
                            min_value=0.0, step=1_000_000.0, format="%.0f")
                    with r2c3:
                        new_funded = st.number_input(
                            "AUM Finance (EUR)",
                            value=float(row_data.get("funded_aum", 0.0)),
                            min_value=0.0, step=1_000_000.0, format="%.0f")

                    r3c1, r3c2, r3c3, r3c4 = st.columns(4)
                    with r3c1:
                        ropts   = [""] + RAISONS_PERTE
                        cur_r   = str(row_data.get("raison_perte") or "")
                        ri      = ropts.index(cur_r) if cur_r in ropts else 0
                        lbl_r   = ("Raison (obligatoire)"
                                   if new_statut in ("Lost","Paused")
                                   else "Raison Perte/Pause")
                        new_raison = st.selectbox(lbl_r, ropts, index=ri)
                    with r3c2:
                        new_concurrent = st.text_input(
                            "Concurrent",
                            value=str(row_data.get("concurrent_choisi") or ""))
                    with r3c3:
                        nad = row_data.get("next_action_date")
                        if not isinstance(nad, date):
                            nad = date.today() + timedelta(days=14)
                        new_nad = st.date_input("Prochaine Action", value=nad)
                    with r3c4:
                        new_sales = st.text_input(
                            "Commercial",
                            value=str(row_data.get("sales_owner") or "Non assigne"))

                    sub = st.form_submit_button(
                        "Sauvegarder les modifications")

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
                            f"Statut : {new_statut}. AUM Finance : {fmt_m(new_funded)}.")
                        st.rerun()
                    else:
                        st.error(msg)

            with st.expander(
                f"Historique des modifications — Deal #{pipeline_id}",
                expanded=False
            ):
                df_audit = db.get_audit_log(pipeline_id)
                if df_audit.empty:
                    st.info("Aucune modification enregistree.")
                else:
                    st.dataframe(
                        df_audit, use_container_width=True, hide_index=True,
                        height=min(240, 45 + len(df_audit) * 36))
    else:
        st.markdown(f"""
        <div style="background:{C_MARINE}04;border:1px dashed {C_MARINE}22;
                    border-radius:8px;padding:22px;text-align:center;margin-top:11px;">
            <div style="color:{C_MARINE};font-weight:600;font-size:0.88rem;">
                Selectionnez un deal dans le tableau pour ouvrir le formulaire</div>
            <div style="color:#888;font-size:0.78rem;margin-top:3px;">
                Le formulaire et l'historique des modifications s'afficheront ici</div>
        </div>""", unsafe_allow_html=True)

    st.divider()
    st.markdown("#### Comparaison AUM par Deal")
    df_viz = df_pipe[
        (df_pipe["target_aum_initial"] > 0) &
        (df_pipe["statut"].isin(["Funded","Soft Commit","Due Diligence","Initial Pitch"]))
    ].copy().head(10)

    if not df_viz.empty:
        fig, ax = plt.subplots(figsize=(12, 3.9))
        fig.patch.set_facecolor("white"); ax.set_facecolor("white")
        x = np.arange(len(df_viz)); w = 0.28
        ax.bar(x-w,   df_viz["target_aum_initial"], w,
               label="AUM Cible",   color=C_GRIS,     edgecolor="white")
        ax.bar(x,     df_viz["revised_aum"],  w,
               label="AUM Revise",  color=C_BLEU_MID, edgecolor="white")
        ax.bar(x+w,   df_viz["funded_aum"],   w,
               label="AUM Finance", color=C_CIEL,     edgecolor="white")
        ax.set_xticks(x)
        ax.set_xticklabels(df_viz["nom_client"].str[:18],
                           rotation=24, ha="right", fontsize=8.5, color=C_MARINE)
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda y,_: f"${y/1e6:.0f}M"))
        ax.tick_params(axis="y", colors=C_MARINE, labelsize=8)
        ax.set_title("AUM Cible / Revise / Finance — Deals Actifs",
                     fontsize=10, fontweight="bold", color=C_MARINE, pad=7)
        ax.legend(fontsize=8.5, frameon=False, labelcolor=C_MARINE)
        ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)
        ax.grid(axis="y", alpha=0.20, color=C_GRIS)
        fig.tight_layout()
        st.pyplot(fig, use_container_width=True)
        plt.close(fig)

    df_lp = df_pipe[df_pipe["statut"].isin(["Lost","Paused"])].copy()
    if not df_lp.empty:
        st.divider()
        st.markdown("#### Deals Perdus / En Pause")
        st.dataframe(
            df_lp[["nom_client","fonds","statut","target_aum_initial",
                   "raison_perte","concurrent_choisi","sales_owner"]],
            use_container_width=True, hide_index=True,
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
# ONGLET 3 : EXECUTIVE DASHBOARD (avec AUM par Region — MISSION 3)
# ============================================================================
with tab_dash:
    st.markdown('<div class="section-title">Executive Dashboard</div>',
                unsafe_allow_html=True)

    kpis = db.get_kpis()

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
    statut_order = [s for s in STATUTS
                    if kpis["statut_repartition"].get(s, 0) > 0]
    if statut_order:
        bcols = st.columns(len(statut_order), gap="small")
        for col, s in zip(bcols, statut_order):
            c_hex = STATUT_COLORS.get(s, C_CIEL)
            with col:
                st.markdown(f"""
                <div style="background:{c_hex}18;border:1px solid {c_hex}44;
                            border-radius:6px;padding:8px;text-align:center;">
                    <div style="font-size:0.64rem;color:{C_MARINE};font-weight:700;
                                text-transform:uppercase;">{s}</div>
                    <div style="font-size:1.45rem;font-weight:800;color:{c_hex};">
                        {kpis['statut_repartition'][s]}</div>
                </div>""", unsafe_allow_html=True)

    st.divider()

    # Alertes
    df_overdue = db.get_overdue_actions()
    if not df_overdue.empty:
        st.markdown(f"""
        <div style="background:#FFF8E1;border-left:4px solid {C_CIEL};
                    border-radius:0 6px 6px 0;padding:10px 14px;margin-bottom:12px;">
            <div style="font-size:0.82rem;font-weight:700;color:{C_MARINE};
                        margin-bottom:6px;">
                {len(df_overdue)} action(s) en retard
            </div>""", unsafe_allow_html=True)
        for _, row in df_overdue.iterrows():
            nad = row.get("next_action_date")
            if isinstance(nad, date):
                days_late = (date.today() - nad).days
                nad_str   = nad.isoformat()
            else:
                days_late = 0; nad_str = str(nad or "—")
            owner = str(row.get("sales_owner","")) or ""
            st.markdown(f"""
            <div class="alert-overdue">
                <b>{row['nom_client']}</b> — {row['fonds']}
                <span style="color:{C_CIEL};font-weight:600;">
                    ({row['statut']})</span>
                — Prevue le <b>{nad_str}</b>
                <span style="color:#B04000;"> ({days_late} j. de retard)</span>
                {f' — <b>{owner}</b>' if owner else ''}
            </div>""", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    # Graphiques : Type Client | Region | AUM par Fonds
    gcol1, gcol2, gcol3 = st.columns([1, 1, 1.2], gap="medium")

    def _donut_ui(ax, lbls, vals, title):
        pal   = [C_CIEL, C_MARINE, C_BLEU_MID, C_BLEU_PALE, C_BLEU_DEEP,
                 "#5BA3C9", "#2C8FBF"]
        cols  = [pal[i % len(pal)] for i in range(len(lbls))]
        _, _, autotexts = ax.pie(
            vals, colors=cols, autopct="%1.1f%%", startangle=90,
            pctdistance=0.74,
            wedgeprops={"width":0.54,"edgecolor":"white","linewidth":1.5}
        )
        for at in autotexts:
            at.set_fontsize(8); at.set_color("white"); at.set_fontweight("bold")
        ax.text(0, 0.07, fmt_m(sum(vals)), ha="center", va="center",
                fontsize=9.5, fontweight="bold", color=C_MARINE)
        ax.text(0,-0.18, "Finance", ha="center", va="center",
                fontsize=7, color="#666")
        patches = [
            matplotlib.patches.Patch(color=cols[i],
                label=f"{lbls[i][:16]}: {fmt_m(vals[i])}")
            for i in range(len(lbls))
        ]
        ax.legend(handles=patches, loc="lower center",
                  bbox_to_anchor=(0.5,-0.26), ncol=2,
                  fontsize=6.5, frameon=False, labelcolor=C_MARINE)
        ax.set_title(title, fontsize=8.5, fontweight="bold",
                     color=C_MARINE, pad=7)

    with gcol1:
        st.markdown("#### Par Type Client")
        if kpis["aum_by_type"]:
            fig1, ax1 = plt.subplots(figsize=(4.2, 4.0))
            fig1.patch.set_facecolor("white")
            _donut_ui(ax1, list(kpis["aum_by_type"].keys()),
                      list(kpis["aum_by_type"].values()), "AUM Finance par Type")
            fig1.tight_layout()
            st.pyplot(fig1, use_container_width=True)
            plt.close(fig1)

    with gcol2:
        st.markdown("#### Par Region")
        aum_reg = db.get_aum_by_region()
        if aum_reg:
            fig2, ax2 = plt.subplots(figsize=(4.2, 4.0))
            fig2.patch.set_facecolor("white")
            _donut_ui(ax2, list(aum_reg.keys()),
                      list(aum_reg.values()), "AUM Finance par Region")
            fig2.tight_layout()
            st.pyplot(fig2, use_container_width=True)
            plt.close(fig2)
        else:
            st.info("Aucun AUM Finance par region.")

    with gcol3:
        st.markdown("#### Par Fonds")
        if kpis["aum_by_fonds"]:
            fig3, ax3 = plt.subplots(figsize=(5.8, 4.0))
            fig3.patch.set_facecolor("white"); ax3.set_facecolor("white")
            flbls = list(kpis["aum_by_fonds"].keys())
            fvals = list(kpis["aum_by_fonds"].values())
            yp    = range(len(flbls))
            bars  = ax3.barh(yp, fvals, color=C_CIEL, edgecolor="white", height=0.50)
            for bar, val in zip(bars, fvals):
                ax3.text(bar.get_width() + max(fvals)*0.01,
                         bar.get_y() + bar.get_height()/2,
                         fmt_m(val), va="center", ha="left",
                         fontsize=8, color=C_MARINE, fontweight="bold")
            ax3.set_yticks(yp)
            ax3.set_yticklabels(flbls, fontsize=8.5, color=C_MARINE)
            ax3.invert_yaxis()
            ax3.xaxis.set_major_formatter(
                plt.FuncFormatter(lambda x,_: fmt_m(x)))
            ax3.tick_params(axis="x", colors=C_MARINE, labelsize=7.5)
            ax3.set_xlim(0, max(fvals)*1.2)
            ax3.spines["top"].set_visible(False)
            ax3.spines["right"].set_visible(False)
            ax3.grid(axis="x", alpha=0.20, color=C_GRIS)
            ax3.set_title("AUM Finance par Fonds",
                          fontsize=8.5, fontweight="bold",
                          color=C_MARINE, pad=7)
            fig3.tight_layout()
            st.pyplot(fig3, use_container_width=True)
            plt.close(fig3)

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
            rank = ["No.1","No.2","No.3"][i] if i < 3 else f"No.{i+1}"
            st.markdown(f"""
            <div style="display:flex;align-items:center;gap:10px;margin:5px 0;
                        padding:8px 12px;background:#F8FAFD;border-radius:6px;
                        border:1px solid {C_MARINE}0C;">
                <div style="font-size:0.76rem;font-weight:700;color:{C_BLEU_MID};
                            min-width:30px;">{rank}</div>
                <div style="flex:1;">
                    <div style="font-size:0.83rem;font-weight:700;
                                color:{C_MARINE};">{row['nom_client']}</div>
                    <div style="font-size:0.7rem;color:#777;">
                        {row['fonds']} &middot; {row['type_client']}
                        &middot; {row['region']}</div>
                    <div style="background:{C_GRIS};border-radius:3px;
                                height:4px;margin-top:4px;overflow:hidden;">
                        <div style="background:{C_CIEL};width:{pct:.0f}%;
                                    height:100%;border-radius:3px;"></div>
                    </div>
                </div>
                <div style="text-align:right;min-width:70px;">
                    <div style="font-size:0.9rem;font-weight:800;color:{C_CIEL};">
                        {fmt_m(val)}</div>
                    <div style="font-size:0.64rem;color:#999;">finance</div>
                </div>
            </div>""", unsafe_allow_html=True)


# ============================================================================
# ONGLET 4 : SALES TRACKING
# ============================================================================
with tab_sales:
    st.markdown('<div class="section-title">Sales Tracking — Suivi par Commercial</div>',
                unsafe_allow_html=True)

    df_sm = db.get_sales_metrics()
    df_na = db.get_next_actions_by_sales(days_ahead=30)

    if df_sm.empty:
        st.info("Aucune donnee de pipeline.")
    else:
        n_own  = len(df_sm)
        s_cols = st.columns(min(n_own, 3), gap="medium")

        for i, (_, row) in enumerate(df_sm.iterrows()):
            with s_cols[i % min(n_own, 3)]:
                retard = (
                    f'<span style="color:#B04000;font-weight:700;">'
                    f'{int(row["Actions en retard"])} en retard</span>'
                    if int(row["Actions en retard"]) > 0
                    else f'<span style="color:{C_CIEL};">A jour</span>'
                )
                st.markdown(f"""
                <div class="sales-card">
                    <div class="sales-card-name">{row['Commercial']}</div>
                    <div style="display:grid;grid-template-columns:1fr 1fr;gap:9px;">
                        <div>
                            <div class="sales-metric">Total Deals</div>
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
                            <div style="font-size:0.75rem;">{retard}</div>
                        </div>
                    </div>
                </div><br>""", unsafe_allow_html=True)

        st.divider()
        st.markdown("#### Tableau de bord commercial")
        df_sd = df_sm.copy()
        for col in ["AUM Finance","Pipeline Actif"]:
            df_sd[col] = df_sd[col].apply(fmt_m)
        st.dataframe(df_sd, use_container_width=True, hide_index=True,
            height=min(340, 50+len(df_sd)*40))

        st.divider()
        st.markdown("#### AUM Finance par Commercial")
        if df_sm["AUM Finance"].sum() > 0:
            fig_s, ax_s = plt.subplots(figsize=(10, 3.4))
            fig_s.patch.set_facecolor("white"); ax_s.set_facecolor("white")
            owners_l = df_sm["Commercial"].tolist()
            aum_v    = df_sm["AUM Finance"].tolist()
            pipe_v   = df_sm["Pipeline Actif"].tolist()
            x = np.arange(len(owners_l)); w = 0.36
            ax_s.bar(x-w/2, aum_v,  w, label="AUM Finance",
                     color=C_CIEL,    edgecolor="white")
            ax_s.bar(x+w/2, pipe_v, w, label="Pipeline Actif",
                     color=C_BLEU_MID, edgecolor="white")
            ax_s.set_xticks(x)
            ax_s.set_xticklabels(owners_l, fontsize=9, color=C_MARINE)
            ax_s.yaxis.set_major_formatter(
                plt.FuncFormatter(lambda y,_: fmt_m(y)))
            ax_s.tick_params(colors=C_MARINE, labelsize=8)
            ax_s.set_title("AUM par Commercial",
                           fontsize=10, fontweight="bold", color=C_MARINE, pad=7)
            ax_s.legend(fontsize=8.5, frameon=False, labelcolor=C_MARINE)
            ax_s.spines["top"].set_visible(False)
            ax_s.spines["right"].set_visible(False)
            ax_s.grid(axis="y", alpha=0.20, color=C_GRIS)
            fig_s.tight_layout()
            st.pyplot(fig_s, use_container_width=True)
            plt.close(fig_s)

        st.divider()
        st.markdown("#### Prochaines Actions — 30 jours")
        if df_na.empty:
            st.info("Aucune action dans les 30 prochains jours.")
        else:
            owners_na = ["Tous"] + sorted(
                df_na["sales_owner"].unique().tolist()
            )
            filter_o = st.selectbox("Filtrer par commercial", owners_na)
            df_nav = (df_na if filter_o == "Tous"
                      else df_na[df_na["sales_owner"] == filter_o])

            today = date.today()
            for _, row in df_nav.iterrows():
                nad = row.get("next_action_date")
                if isinstance(nad, date):
                    delta = (nad - today).days
                    if delta < 0:
                        timing = f"RETARD ({abs(delta)} j.)"; dot = "#B04000"
                    elif delta == 0:
                        timing = "Aujourd'hui"; dot = C_BLEU_MID
                    elif delta <= 7:
                        timing = f"Dans {delta} j."; dot = C_CIEL
                    else:
                        timing = f"Dans {delta} j."; dot = C_MARINE
                    nad_s = nad.isoformat()
                else:
                    timing = "—"; dot = C_GRIS; nad_s = "—"

                revised = float(row.get("revised_aum", 0) or 0)
                st.markdown(f"""
                <div style="display:flex;align-items:center;gap:11px;
                            padding:7px 11px;margin:3px 0;
                            background:#F8FAFD;border-radius:5px;
                            border-left:3px solid {dot};">
                    <div style="min-width:105px;font-size:0.74rem;
                                color:{dot};font-weight:700;">{timing}</div>
                    <div style="flex:1;">
                        <span style="font-weight:600;color:{C_MARINE};
                                     font-size:0.82rem;">{row['nom_client']}</span>
                        <span style="color:#888;font-size:0.74rem;">
                            &nbsp;{row['fonds']} &middot; {row['statut']}</span>
                    </div>
                    <div style="font-size:0.78rem;color:{C_BLEU_MID};
                                font-weight:600;min-width:68px;text-align:right;">
                        {fmt_m(revised)}</div>
                    <div style="font-size:0.74rem;color:#888;min-width:88px;
                                text-align:right;">{row['sales_owner']}</div>
                </div>""", unsafe_allow_html=True)


# ============================================================================
# ONGLET 5 : PERFORMANCE ET NAV (MISSION 4)
# ============================================================================
with tab_perf:
    st.markdown('<div class="section-title">Performance et NAV</div>',
                unsafe_allow_html=True)

    col_up, col_info = st.columns([1.2, 1], gap="large")

    with col_info:
        st.markdown(f"""
        <div style="background:{C_MARINE}06;border:1px solid {C_MARINE}12;
                    border-radius:8px;padding:15px 17px;">
            <div style="font-weight:700;color:{C_MARINE};font-size:0.9rem;
                        margin-bottom:9px;">Format attendu</div>
            <div style="font-size:0.79rem;color:#444;line-height:1.75;">
                <b>Colonnes obligatoires</b><br>
                &bull; <code>Date</code> — YYYY-MM-DD ou DD/MM/YYYY<br>
                &bull; <code>Fonds</code> — Nom du fonds<br>
                &bull; <code>NAV</code> — Valeur liquidative numerique<br><br>
                <b>Calculs produits</b><br>
                &bull; Base 100 normalisee (protection division par zero)<br>
                &bull; Performance 1 Mois glissant<br>
                &bull; Performance YTD<br>
                &bull; Performance periode complete<br><br>
                <b>Integration PDF</b><br>
                Les donnees NAV chargees sont automatiquement
                integrees dans l'export PDF, filtrées sur le
                perimetre selectionne dans la barre laterale.
            </div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("#### Demonstration")
        if st.button("Generer un fichier NAV de demonstration",
                     use_container_width=True):
            demo_dates = pd.date_range(
                f"{date.today().year-1}-01-01", date.today(), freq="B"
            )
            rng  = np.random.default_rng(42)
            navs = {f: 100.0 for f in FONDS}
            rows = []
            for d in demo_dates:
                for fonds, nav in navs.items():
                    nav *= (1 + rng.normal(0.0003, 0.006))
                    navs[fonds] = nav
                    rows.append({"Date": d.date().isoformat(),
                                 "Fonds": fonds, "NAV": round(nav, 4)})
            buf_demo = io.BytesIO()
            pd.DataFrame(rows).to_excel(buf_demo, index=False)
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
                         f"Trouvees : {list(df_nav.columns)}")
                st.stop()

            # Nettoyage securise — format mixed, dayfirst, drop NaT immediatement
            df_nav["Date"] = pd.to_datetime(
                df_nav["Date"], format="mixed", dayfirst=True, errors="coerce"
            )
            df_nav = df_nav.dropna(subset=["Date"])
            df_nav["NAV"]   = pd.to_numeric(df_nav["NAV"], errors="coerce")
            df_nav = df_nav.dropna(subset=["NAV"])
            df_nav["Fonds"] = df_nav["Fonds"].astype(str).str.strip()
            df_nav = df_nav.sort_values("Date").reset_index(drop=True)

            if df_nav.empty:
                st.error("Aucune donnee valide apres nettoyage.")
                st.stop()

            fonds_list = sorted(df_nav["Fonds"].unique().tolist())
            d_min = df_nav["Date"].min(); d_max = df_nav["Date"].max()

            st.markdown(f"""
            <div style="background:{C_CIEL}14;border-left:3px solid {C_CIEL};
                        border-radius:0 6px 6px 0;padding:8px 14px;margin:9px 0;">
                {len(df_nav):,} points &mdash; {len(fonds_list)} fonds &mdash;
                {d_min.strftime('%d/%m/%Y')} au {d_max.strftime('%d/%m/%Y')}
            </div>""", unsafe_allow_html=True)

            # Filtres periode & fonds
            st.markdown("---")
            ff1, ff2, ff3 = st.columns([2, 1, 1])
            with ff1:
                fonds_sel = st.multiselect(
                    "Fonds a afficher", fonds_list,
                    default=fonds_list[:min(5, len(fonds_list))]
                )
            with ff2:
                d_debut = st.date_input("Depuis", value=d_min.date())
            with ff3:
                d_fin   = st.date_input("Jusqu'au", value=d_max.date())

            if not fonds_sel:
                st.warning("Selectionnez au moins un fonds.")
                st.stop()

            mask = (
                df_nav["Fonds"].isin(fonds_sel) &
                (df_nav["Date"].dt.date >= d_debut) &
                (df_nav["Date"].dt.date <= d_fin)
            )
            df_fn = df_nav[mask].copy()
            if df_fn.empty:
                st.warning("Aucune donnee pour la periode selectionnee.")
                st.stop()

            pivot = (df_fn
                     .pivot_table(index="Date", columns="Fonds",
                                  values="NAV", aggfunc="last")
                     .sort_index().ffill())
            pivot = pivot[[f for f in fonds_sel if f in pivot.columns]]

            # Base 100 securisee (protection division par zero)
            base100 = pivot.copy() * np.nan
            for fonds in pivot.columns:
                s = pivot[fonds].dropna()
                if s.empty: continue
                first_val = float(s.iloc[0])
                if first_val != 0 and not np.isnan(first_val):
                    base100[fonds] = pivot[fonds] / first_val * 100

            # Stockage session_state pour le PDF
            st.session_state["nav_base100"] = base100
            st.session_state["perf_fonds"]  = fonds_sel

            # --- Graphique Base 100 (matplotlib, fond blanc, minimaliste) ---
            st.markdown("#### Evolution NAV — Base 100")

            fig_nav, ax_nav = plt.subplots(figsize=(12, 5))
            fig_nav.patch.set_facecolor("white")
            ax_nav.set_facecolor("white")

            plotted = 0
            for i, fonds in enumerate(pivot.columns):
                series = base100[fonds].dropna()
                if series.empty: continue
                color    = NAV_PALETTE[i % len(NAV_PALETTE)]
                line_sty = "-" if i % 2 == 0 else "--"

                # Protection single-point
                if len(series) >= 2:
                    ax_nav.plot(series.index, series.values,
                                label=fonds, color=color,
                                linewidth=1.8, linestyle=line_sty, alpha=0.9)
                else:
                    ax_nav.scatter(series.index, series.values,
                                   color=color, s=60, label=fonds, zorder=5)
                if not series.empty:
                    ax_nav.scatter([series.index[-1]], [series.values[-1]],
                                   color=color, s=32, zorder=6)
                plotted += 1

            if plotted == 0:
                st.warning("Aucune Base 100 calculable.")
                st.stop()

            ax_nav.axhline(100, color=C_GRIS, linewidth=0.8, linestyle=":")
            ax_nav.set_ylabel("NAV (Base 100)", fontsize=8.5, color=C_MARINE)
            ax_nav.tick_params(colors=C_MARINE, labelsize=8)
            # Minimalisme — pas de grilles verticales
            ax_nav.spines["top"].set_visible(False)
            ax_nav.spines["right"].set_visible(False)
            ax_nav.spines["left"].set_color(C_GRIS)
            ax_nav.spines["bottom"].set_color(C_GRIS)
            ax_nav.grid(axis="y", alpha=0.18, color=C_GRIS, linewidth=0.6)
            ax_nav.legend(fontsize=8.5, frameon=True, framealpha=0.93,
                          edgecolor=C_GRIS, labelcolor=C_MARINE, loc="upper left")
            ax_nav.set_title(
                f"Performance NAV — Base 100 — "
                f"{d_debut.strftime('%d/%m/%Y')} au {d_fin.strftime('%d/%m/%Y')}",
                fontsize=10.5, fontweight="bold", color=C_MARINE, pad=9
            )
            plt.xticks(rotation=16, ha="right", fontsize=7.5)
            fig_nav.tight_layout()
            st.pyplot(fig_nav, use_container_width=True)
            plt.close(fig_nav)

            # --- Calcul performances ---
            today_ts  = pd.Timestamp(date.today())
            one_m_ago = today_ts - pd.DateOffset(months=1)
            jan_1     = pd.Timestamp(f"{date.today().year}-01-01")

            perf_rows = []
            for fonds in pivot.columns:
                series = pivot[fonds].dropna()
                if series.empty: continue
                nav_last  = float(series.iloc[-1])
                nav_first = float(series.iloc[0])

                s_1m  = series[series.index >= one_m_ago]
                p1m   = ((nav_last / float(s_1m.iloc[0]) - 1) * 100
                         if len(s_1m) > 0 and float(s_1m.iloc[0]) != 0
                         else float("nan"))

                s_ytd  = series[series.index >= jan_1]
                pytd   = ((nav_last / float(s_ytd.iloc[0]) - 1) * 100
                          if len(s_ytd) > 0 and float(s_ytd.iloc[0]) != 0
                          else float("nan"))

                pp = ((nav_last / nav_first - 1) * 100
                      if nav_first != 0 else float("nan"))

                b100s  = base100[fonds].dropna()
                nb100  = (float(b100s.iloc[-1]) if not b100s.empty
                          else float("nan"))

                perf_rows.append({
                    "Fonds":            fonds,
                    "NAV Derniere":     round(nav_last, 4),
                    "Base 100 Actuel":  (round(nb100, 2)
                                         if not np.isnan(nb100) else None),
                    "Perf 1M (%)":      (round(p1m, 2)
                                         if not np.isnan(p1m) else None),
                    "Perf YTD (%)":     (round(pytd, 2)
                                         if not np.isnan(pytd) else None),
                    "Perf Periode (%)": (round(pp, 2)
                                         if not np.isnan(pp) else None),
                })

            if perf_rows:
                df_pt = pd.DataFrame(perf_rows)
                # Stockage pour PDF
                st.session_state["perf_data"] = df_pt

                # Tableau HTML des performances
                st.markdown("#### Tableau des Performances")

                def _fp(val):
                    if val is None or (isinstance(val, float) and np.isnan(val)):
                        return f'<span style="color:#999;">n.d.</span>'
                    c = C_CIEL if val >= 0 else "#8B2020"
                    s = "+" if val > 0 else ""
                    return (f'<span style="color:{c};font-weight:700;">'
                            f'{s}{val:.2f}%</span>')

                tbl_h = f"""
                <table style="width:100%;border-collapse:collapse;font-size:0.82rem;">
                    <thead><tr style="background:{C_MARINE};color:white;">
                        <th style="padding:8px 12px;text-align:left;">Fonds</th>
                        <th style="padding:8px 12px;text-align:right;">NAV</th>
                        <th style="padding:8px 12px;text-align:right;">Base 100</th>
                        <th style="padding:8px 12px;text-align:right;">Perf 1M</th>
                        <th style="padding:8px 12px;text-align:right;">Perf YTD</th>
                        <th style="padding:8px 12px;text-align:right;">Perf Periode</th>
                    </tr></thead><tbody>"""
                for i, r in enumerate(perf_rows):
                    bg = "#F8FAFD" if i % 2 == 0 else C_BLANC
                    tbl_h += f"""
                    <tr style="background:{bg};border-bottom:1px solid {C_GRIS};">
                        <td style="padding:7px 12px;font-weight:600;
                                   color:{C_MARINE};">{r['Fonds']}</td>
                        <td style="padding:7px 12px;text-align:right;
                                   color:{C_MARINE};">{r['NAV Derniere']:.4f}</td>
                        <td style="padding:7px 12px;text-align:right;
                                   color:{C_MARINE};">
                            {f"{r['Base 100 Actuel']:.2f}" if r['Base 100 Actuel'] else 'n.d.'}</td>
                        <td style="padding:7px 12px;text-align:right;">
                            {_fp(r['Perf 1M (%)'])}</td>
                        <td style="padding:7px 12px;text-align:right;">
                            {_fp(r['Perf YTD (%)'])}</td>
                        <td style="padding:7px 12px;text-align:right;">
                            {_fp(r['Perf Periode (%)'])}</td>
                    </tr>"""
                tbl_h += "</tbody></table>"
                st.markdown(tbl_h, unsafe_allow_html=True)
                st.markdown("<br>", unsafe_allow_html=True)

                st.download_button(
                    "Exporter le tableau en CSV",
                    data=df_pt.to_csv(index=False).encode("utf-8"),
                    file_name=f"performances_{date.today().isoformat()}.csv",
                    mime="text/csv"
                )

                # Graphique YTD — fond blanc, minimaliste (MISSION 4)
                ytd_data = [r for r in perf_rows if r["Perf YTD (%)"] is not None]
                if ytd_data:
                    ytd_data.sort(key=lambda r: r["Perf YTD (%)"], reverse=True)
                    st.markdown("#### Comparaison Performances YTD")

                    fig_ytd, ax_ytd = plt.subplots(figsize=(10, 3.3))
                    fig_ytd.patch.set_facecolor("white")
                    ax_ytd.set_facecolor("white")

                    fonds_y = [r["Fonds"] for r in ytd_data]
                    vals_y  = [r["Perf YTD (%)"] for r in ytd_data]
                    bar_c   = [C_CIEL if v >= 0 else C_GRIS for v in vals_y]

                    bars_y = ax_ytd.bar(
                        range(len(fonds_y)), vals_y,
                        color=bar_c, edgecolor="white",
                        linewidth=0.4, width=0.6
                    )
                    ax_ytd.axhline(0, color=C_MARINE, linewidth=0.6)
                    for bar, val in zip(bars_y, vals_y):
                        sign = "+" if val > 0 else ""
                        ypos = val + 0.04 if val >= 0 else val - 0.22
                        ax_ytd.text(
                            bar.get_x() + bar.get_width()/2, ypos,
                            f"{sign}{val:.2f}%",
                            ha="center",
                            va="bottom" if val >= 0 else "top",
                            fontsize=8, color=C_MARINE, fontweight="bold"
                        )
                    ax_ytd.set_xticks(range(len(fonds_y)))
                    ax_ytd.set_xticklabels(fonds_y, rotation=14, ha="right",
                                           fontsize=8.5, color=C_MARINE)
                    ax_ytd.set_ylabel("YTD (%)", fontsize=8.5, color=C_MARINE)
                    ax_ytd.tick_params(colors=C_MARINE, labelsize=8)
                    # Minimalisme : pas de grille verticale
                    ax_ytd.spines["top"].set_visible(False)
                    ax_ytd.spines["right"].set_visible(False)
                    ax_ytd.spines["left"].set_color(C_GRIS)
                    ax_ytd.spines["bottom"].set_color(C_GRIS)
                    ax_ytd.grid(axis="y", alpha=0.18, color=C_GRIS)
                    ax_ytd.set_title("Performance YTD par Fonds (%)",
                                     fontsize=10, fontweight="bold",
                                     color=C_MARINE, pad=7)
                    fig_ytd.tight_layout()
                    st.pyplot(fig_ytd, use_container_width=True)
                    plt.close(fig_ytd)

        except Exception as e:
            st.error(f"Erreur traitement fichier NAV : {e}")
            import traceback
            with st.expander("Details"):
                st.code(traceback.format_exc())

    else:
        st.markdown(f"""
        <div style="background:{C_MARINE}04;border:2px dashed {C_MARINE}1E;
                    border-radius:10px;padding:44px;text-align:center;margin-top:12px;">
            <div style="font-size:0.96rem;font-weight:700;color:{C_MARINE};
                        margin-bottom:5px;">Module Performance et NAV</div>
            <div style="color:#777;font-size:0.81rem;max-width:380px;
                        margin:0 auto;line-height:1.65;">
                Chargez un fichier Excel ou CSV avec les colonnes
                <code>Date</code>, <code>Fonds</code>, <code>NAV</code>
                pour generer les courbes Base 100 et le tableau de
                performances.<br><br>
                Utilisez le bouton <b>Generer un fichier NAV de demonstration</b>
                pour un exemple testable immediatement.
            </div>
        </div>""", unsafe_allow_html=True)
