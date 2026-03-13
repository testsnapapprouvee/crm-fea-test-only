# =============================================================================
# app.py — CRM & Reporting Tool — Asset Management
# Refactoring Staff Engineer : Master-Detail Pipeline + Module Performance NAV
# Charte Amundi : #002D54 (Marine) | #00A8E1 (Ciel)
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
REGIONS           = ["GCC", "EMEA", "APAC", "Nordics", "Asia ex-Japan", "North America", "LatAm"]
FONDS             = ["Global Value", "International Fund", "Income Builder",
                     "Resilient Equity", "Private Debt", "Active ETFs"]
STATUTS           = ["Prospect", "Initial Pitch", "Due Diligence",
                     "Soft Commit", "Funded", "Paused", "Lost", "Redeemed"]
STATUTS_ACTIFS    = ["Prospect", "Initial Pitch", "Due Diligence", "Soft Commit"]
RAISONS_PERTE     = ["Pricing", "Track Record", "Macro", "Competitor", "Autre"]
TYPES_INTERACTION = ["Call", "Meeting", "Email", "Roadshow", "Conference", "Autre"]

# Palette NAV — strictement bleus Amundi, jamais couleurs Matplotlib par défaut
NAV_PALETTE = [C_CIEL, C_MARINE, C_BLEU_MID, C_BLEU_PALE, C_BLEU_DEEP,
               "#5BA3C9", "#2C8FBF", "#004F8C", "#A8D8EE", "#003060"]


# ---------------------------------------------------------------------------
# CONFIG STREAMLIT
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="CRM Asset Management",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ---------------------------------------------------------------------------
# CSS — CHARTE AMUNDI STRICTE
# ---------------------------------------------------------------------------
st.markdown(f"""
<style>
    .stApp, .main .block-container {{
        background-color: {C_BLANC};
        color: {C_MARINE};
    }}
    [data-testid="stSidebar"] {{ background-color: {C_MARINE}; }}
    [data-testid="stSidebar"] * {{ color: {C_BLANC} !important; }}
    [data-testid="stSidebar"] h1,[data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3 {{ color: {C_CIEL} !important; }}

    .crm-header {{
        background: linear-gradient(135deg, {C_MARINE} 0%, #003D6B 100%);
        color: {C_BLANC};
        padding: 22px 28px;
        border-radius: 10px;
        margin-bottom: 20px;
        border-left: 5px solid {C_CIEL};
    }}
    .crm-header h1 {{ color: {C_BLANC} !important; margin:0;
                      font-size:1.75rem; font-weight:700; letter-spacing:-0.3px; }}
    .crm-header p  {{ color: {C_CIEL}; margin:4px 0 0 0; font-size:0.87rem; }}

    .kpi-card {{
        background: linear-gradient(135deg, {C_MARINE} 0%, #004070 100%);
        padding: 18px 14px;
        border-radius: 10px;
        text-align: center;
        border: 1px solid {C_CIEL}33;
    }}
    .kpi-label {{ font-size:0.73rem; color:{C_CIEL}; text-transform:uppercase;
                  letter-spacing:0.9px; margin-bottom:8px; font-weight:600; }}
    .kpi-value {{ font-size:1.65rem; font-weight:800; color:{C_BLANC}; line-height:1.1; }}
    .kpi-sub   {{ font-size:0.7rem; color:{C_GRIS}; margin-top:4px; }}

    .section-title {{
        font-size:1.05rem; font-weight:700; color:{C_MARINE};
        border-bottom:3px solid {C_CIEL}; padding-bottom:6px; margin:18px 0 12px 0;
    }}

    /* ---- Master-Detail Panel ---- */
    .detail-panel {{
        background: linear-gradient(135deg, #F0F6FC 0%, #E8F2FA 100%);
        border: 1.5px solid {C_CIEL}55;
        border-radius: 12px;
        padding: 22px 24px 18px 24px;
        margin-top: 18px;
        box-shadow: 0 4px 20px {C_MARINE}12;
    }}
    .detail-panel-title {{
        font-size: 1rem; font-weight: 700; color: {C_MARINE};
        margin-bottom: 16px; display: flex; align-items: center; gap: 8px;
    }}
    .detail-badge {{
        display:inline-block; padding:3px 12px; border-radius:12px;
        font-size:0.78rem; font-weight:600;
    }}

    /* ---- Pipeline read-only table ---- */
    .pipeline-hint {{
        background: {C_CIEL}12; border-left: 3px solid {C_CIEL};
        border-radius: 0 6px 6px 0; padding: 8px 14px;
        font-size: 0.83rem; color: {C_MARINE}; margin-bottom: 12px;
    }}

    /* ---- Alertes ---- */
    .alert-overdue {{
        background:#FFF8E1; border-left:4px solid {C_CIEL};
        border-radius:0 6px 6px 0; padding:9px 14px; margin:5px 0;
        font-size:0.83rem; color:{C_MARINE};
    }}

    /* ---- Onglets ---- */
    .stTabs [data-baseweb="tab-list"] {{
        background:{C_BLANC}; border-bottom:2px solid {C_MARINE}20; gap:4px;
    }}
    .stTabs [data-baseweb="tab"] {{
        color:{C_MARINE}; font-weight:600; font-size:0.87rem;
        padding:10px 20px; border-radius:6px 6px 0 0;
        background:{C_MARINE}0A;
    }}
    .stTabs [aria-selected="true"] {{
        background:{C_MARINE} !important; color:{C_BLANC} !important;
    }}

    /* ---- Boutons ---- */
    .stButton > button {{
        background:{C_MARINE}; color:{C_BLANC}; border:none;
        border-radius:6px; font-weight:600; padding:8px 20px;
        transition: all 0.2s;
    }}
    .stButton > button:hover {{
        background:{C_CIEL}; color:{C_BLANC};
        box-shadow:0 4px 12px {C_CIEL}44;
    }}

    /* ---- Labels formulaires ---- */
    .stSelectbox label, .stTextInput label, .stNumberInput label,
    .stDateInput label, .stTextArea label {{
        color:{C_MARINE} !important; font-weight:600; font-size:0.82rem;
    }}

    /* ---- Performance module ---- */
    .perf-positive {{ color: {C_CIEL}; font-weight:700; }}
    .perf-negative {{ color: #8B2020; font-weight:700; }}

    h1,h2,h3,h4 {{ color:{C_MARINE} !important; }}
    hr {{ border-color:{C_MARINE}18; }}
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
    """Formate un montant AUM pour l'UI."""
    if v >= 1_000_000_000: return f"${v/1_000_000_000:.2f}Md"
    if v >= 1_000_000:     return f"${v/1_000_000:.1f}M"
    if v >= 1_000:         return f"${v/1_000:.0f}k"
    return f"${v:.0f}"


def statut_badge(statut: str) -> str:
    colors = {
        "Funded":        (C_CIEL,     C_BLANC),
        "Soft Commit":   (C_BLEU_MID, C_BLANC),
        "Due Diligence": ("#004F8C",  C_BLANC),
        "Initial Pitch": ("#3A7EBA",  C_BLANC),
        "Prospect":      (f"{C_MARINE}22", C_MARINE),
        "Lost":          (C_GRIS,    "#555"),
        "Paused":        ("#C8D8E8",  C_MARINE),
        "Redeemed":      ("#D0E4F0",  C_MARINE),
    }
    bg, fg = colors.get(statut, (C_GRIS, "#555"))
    return (f'<span class="detail-badge" '
            f'style="background:{bg};color:{fg};">{statut}</span>')


def kpi_card(label, value, sub=""):
    return f"""
    <div class="kpi-card">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{value}</div>
        {"<div class='kpi-sub'>"+sub+"</div>" if sub else ""}
    </div>"""


# ---------------------------------------------------------------------------
# SIDEBAR
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown(f"""
    <div style="text-align:center;padding:10px 0 20px 0;">
        <div style="font-size:2.1rem;">📊</div>
        <div style="font-size:1.05rem;font-weight:800;color:{C_CIEL};">
            CRM Asset Management
        </div>
        <div style="font-size:0.73rem;color:{C_GRIS};margin-top:4px;">
            {date.today().strftime("%d %B %Y")}
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.divider()

    kpis_side = db.get_kpis()
    st.markdown(f'<div style="color:{C_CIEL};font-size:0.76rem;font-weight:700;'
                f'text-transform:uppercase;letter-spacing:1px;margin-bottom:8px;">'
                f'📈 Snapshot</div>', unsafe_allow_html=True)
    for label, val in [
        ("AUM Financé Total", fmt_m(kpis_side["total_funded"])),
        ("Pipeline Actif",    fmt_m(kpis_side["pipeline_actif"])),
    ]:
        st.markdown(f"""
        <div style="background:{C_BLANC}15;padding:11px;border-radius:8px;margin-bottom:7px;">
            <div style="font-size:0.7rem;color:{C_GRIS};">{label}</div>
            <div style="font-size:1.25rem;font-weight:800;color:{C_CIEL};">{val}</div>
        </div>""", unsafe_allow_html=True)

    st.divider()
    st.markdown(f'<div style="color:{C_CIEL};font-size:0.76rem;font-weight:700;'
                f'text-transform:uppercase;letter-spacing:1px;margin-bottom:8px;">'
                f'📄 Export Report</div>', unsafe_allow_html=True)

    mode_comex = st.toggle("🔒 Mode Comex (Anonymisation)", value=False,
        help="Remplace les noms clients par {Type} – {Région} dans le PDF.")
    if mode_comex:
        st.caption("⚠️ Anonymisation activée.")

    if st.button("⬇️ Exporter Executive Report", use_container_width=True):
        with st.spinner("Génération du rapport PDF..."):
            try:
                pdf_bytes = pdf_gen.generate_pdf(
                    pipeline_df=db.get_pipeline_with_clients(),
                    kpis=db.get_kpis(),
                    mode_comex=mode_comex
                )
                fname = (f"report_comex_{date.today().isoformat()}.pdf"
                         if mode_comex else f"report_{date.today().isoformat()}.pdf")
                st.download_button("📥 Télécharger le PDF", data=pdf_bytes,
                                   file_name=fname, mime="application/pdf",
                                   use_container_width=True)
                st.success("✅ Rapport généré.")
            except Exception as e:
                st.error(f"❌ Erreur : {e}")

    st.divider()
    st.caption("Version 2.0 — Charte Amundi")


# ---------------------------------------------------------------------------
# EN-TÊTE
# ---------------------------------------------------------------------------
st.markdown(f"""
<div class="crm-header">
    <h1>📊 CRM & Reporting — Asset Management</h1>
    <p>Gestion du pipeline commercial · Suivi des mandats · Reporting Exécutif · Performance NAV</p>
</div>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# ONGLETS
# ---------------------------------------------------------------------------
tab_ingest, tab_pipeline, tab_dash, tab_perf = st.tabs([
    "📥  Data Ingestion",
    "🔄  Pipeline Management",
    "📊  Executive Dashboard",
    "📈  Performance & NAV",
])


# ============================================================================
# ONGLET 1 : DATA INGESTION  (inchangé — robuste)
# ============================================================================
with tab_ingest:
    st.markdown('<div class="section-title">📥 Saisie & Import de Données</div>',
                unsafe_allow_html=True)

    col_form, col_import = st.columns([1, 1], gap="large")

    with col_form:
        with st.expander("➕ Ajouter un Client", expanded=True):
            with st.form("form_add_client", clear_on_submit=True):
                c1, c2 = st.columns(2)
                with c1:
                    nom_client  = st.text_input("Nom du Client *")
                    type_client = st.selectbox("Type Client *", TYPES_CLIENT)
                with c2:
                    region = st.selectbox("Région *", REGIONS)
                sub_c = st.form_submit_button("✅ Enregistrer", use_container_width=True)
            if sub_c:
                if not nom_client.strip():
                    st.error("⚠️ Nom obligatoire.")
                else:
                    try:
                        db.add_client(nom_client.strip(), type_client, region)
                        st.success(f"✅ Client **{nom_client}** ajouté.")
                    except Exception as e:
                        st.warning("⚠️ Ce client existe déjà." if "UNIQUE" in str(e) else f"❌ {e}")

        with st.expander("➕ Ajouter un Deal Pipeline", expanded=True):
            clients_dict = db.get_client_options()
            if not clients_dict:
                st.info("Ajoutez d'abord un client.")
            else:
                with st.form("form_add_deal", clear_on_submit=True):
                    ca, cb = st.columns(2)
                    with ca:
                        client_sel = st.selectbox("Client *", list(clients_dict.keys()))
                        fonds_sel  = st.selectbox("Fonds *", FONDS)
                        statut_sel = st.selectbox("Statut *", STATUTS)
                    with cb:
                        target_aum  = st.number_input("AUM Cible (€)", min_value=0, step=1_000_000)
                        revised_aum = st.number_input("AUM Révisé (€)", min_value=0, step=1_000_000)
                        funded_aum  = st.number_input("AUM Financé (€)", min_value=0, step=1_000_000)

                    raison_perte, concurrent = "", ""
                    if statut_sel in ("Lost", "Paused"):
                        cc, cd = st.columns(2)
                        with cc: raison_perte = st.selectbox("Raison *", RAISONS_PERTE)
                        with cd: concurrent   = st.text_input("Concurrent")

                    next_action = st.date_input("Prochaine Action",
                                                value=date.today() + timedelta(days=14))
                    sub_d = st.form_submit_button("✅ Enregistrer le Deal", use_container_width=True)

                if sub_d:
                    if statut_sel in ("Lost", "Paused") and not raison_perte:
                        st.error("⚠️ Raison obligatoire.")
                    else:
                        db.add_pipeline_entry(
                            clients_dict[client_sel], fonds_sel, statut_sel,
                            float(target_aum), float(revised_aum), float(funded_aum),
                            raison_perte, concurrent, next_action.isoformat()
                        )
                        st.success(f"✅ Deal **{fonds_sel}** / **{client_sel}** ajouté.")

        with st.expander("📝 Enregistrer une Activité"):
            clients_dict2 = db.get_client_options()
            if clients_dict2:
                with st.form("form_act", clear_on_submit=True):
                    ce, cf = st.columns(2)
                    with ce:
                        act_client = st.selectbox("Client *", list(clients_dict2.keys()))
                        act_type   = st.selectbox("Type", TYPES_INTERACTION)
                    with cf:
                        act_date  = st.date_input("Date", value=date.today())
                        act_notes = st.text_area("Notes", height=80)
                    sub_a = st.form_submit_button("✅ Enregistrer", use_container_width=True)
                if sub_a:
                    db.add_activity(clients_dict2[act_client],
                                    act_date.isoformat(), act_notes, act_type)
                    st.success(f"✅ Activité enregistrée pour **{act_client}**.")

    with col_import:
        st.markdown("#### 📂 Import CSV / Excel (Upsert)")
        import_type    = st.radio("Table cible", ["Clients", "Pipeline"], horizontal=True)
        uploaded_file  = st.file_uploader("Fichier CSV ou Excel (.xlsx)", type=["csv","xlsx","xls"])

        if import_type == "Clients":
            st.info("📋 Colonnes : `nom_client`, `type_client`, `region`")
        else:
            st.info("📋 Colonnes : `nom_client`, `fonds`, `statut`, `target_aum_initial`, "
                    "`revised_aum`, `funded_aum`, `raison_perte`, `concurrent_choisi`, `next_action_date`")

        if uploaded_file:
            try:
                df_imp = (pd.read_csv(uploaded_file) if uploaded_file.name.endswith(".csv")
                          else pd.read_excel(uploaded_file))
                st.dataframe(df_imp.head(5), use_container_width=True, height=160)
                st.caption(f"{len(df_imp)} ligne(s)")
                if st.button("🔄 Lancer l'Upsert", use_container_width=True):
                    with st.spinner("Import..."):
                        fn = db.upsert_clients_from_df if import_type == "Clients" \
                             else db.upsert_pipeline_from_df
                        ins, upd = fn(df_imp)
                    st.success(f"✅ {ins} créé(s), {upd} mis à jour.")
            except Exception as e:
                st.error(f"❌ {e}")

        st.divider()
        st.markdown("#### 📅 Dernières Activités")
        df_act = db.get_activities()
        if not df_act.empty:
            st.dataframe(
                df_act[["nom_client","date","type_interaction","notes"]].head(10),
                use_container_width=True, height=270, hide_index=True,
                column_config={
                    "nom_client":       st.column_config.TextColumn("Client"),
                    "date":             st.column_config.TextColumn("Date"),
                    "type_interaction": st.column_config.TextColumn("Type"),
                    "notes":            st.column_config.TextColumn("Notes"),
                }
            )


# ============================================================================
# ONGLET 2 : PIPELINE MANAGEMENT — MASTER-DETAIL
# ============================================================================
with tab_pipeline:
    st.markdown('<div class="section-title">🔄 Pipeline Management</div>',
                unsafe_allow_html=True)

    # ---- Filtres ----
    with st.expander("🔍 Filtres", expanded=False):
        fc1, fc2, fc3 = st.columns(3)
        with fc1:
            filt_statuts = st.multiselect("Statuts", STATUTS, default=STATUTS_ACTIFS)
        with fc2:
            filt_fonds   = st.multiselect("Fonds", FONDS)
        with fc3:
            filt_regions = st.multiselect("Régions", REGIONS)

    # ---- Chargement & filtrage ----
    df_pipe = db.get_pipeline_with_clients()
    df_view = df_pipe.copy()
    if filt_statuts:  df_view = df_view[df_view["statut"].isin(filt_statuts)]
    if filt_fonds:    df_view = df_view[df_view["fonds"].isin(filt_fonds)]
    if filt_regions:  df_view = df_view[df_view["region"].isin(filt_regions)]

    st.markdown(
        f'<div class="pipeline-hint">👆 Cliquez sur une ligne pour ouvrir le formulaire '
        f'de modification — <b>{len(df_view)} deal(s)</b> affichés</div>',
        unsafe_allow_html=True
    )

    # ---- Préparer la vue affichée (colonne date → str pour l'affichage) ----
    df_display = df_view.copy()
    df_display["next_action_date"] = df_display["next_action_date"].apply(
        lambda d: d.isoformat() if isinstance(d, date) else ""
    )
    # Colonnes d'affichage (pas client_id, pas les AUM bruts)
    cols_show = ["id", "nom_client", "type_client", "region", "fonds", "statut",
                 "target_aum_initial", "revised_aum", "funded_aum",
                 "raison_perte", "concurrent_choisi", "next_action_date"]

    # ---- DataFrame en LECTURE SEULE avec sélection de ligne ----
    event = st.dataframe(
        df_display[cols_show],
        use_container_width=True,
        height=390,
        hide_index=True,
        on_select="rerun",
        selection_mode="single-row",
        column_config={
            "id":                  st.column_config.NumberColumn("ID", width="small"),
            "nom_client":          st.column_config.TextColumn("Client", width="medium"),
            "type_client":         st.column_config.TextColumn("Type", width="small"),
            "region":              st.column_config.TextColumn("Région", width="small"),
            "fonds":               st.column_config.TextColumn("Fonds", width="medium"),
            "statut":              st.column_config.TextColumn("Statut", width="small"),
            "target_aum_initial":  st.column_config.NumberColumn("AUM Cible", format="€ %,.0f"),
            "revised_aum":         st.column_config.NumberColumn("AUM Révisé", format="€ %,.0f"),
            "funded_aum":          st.column_config.NumberColumn("AUM Financé", format="€ %,.0f"),
            "raison_perte":        st.column_config.TextColumn("Raison", width="small"),
            "concurrent_choisi":   st.column_config.TextColumn("Concurrent", width="small"),
            "next_action_date":    st.column_config.TextColumn("Next Action", width="small"),
        },
        key="pipeline_readonly"
    )

    # ---- PANNEAU MASTER-DETAIL ----
    selected_rows = event.selection.rows if event.selection else []

    if len(selected_rows) > 0:
        selected_idx  = selected_rows[0]
        selected_row  = df_view.iloc[selected_idx]
        pipeline_id   = int(selected_row["id"])

        # Récupérer les données fraîches depuis la DB (types garantis)
        row_data = db.get_pipeline_row_by_id(pipeline_id)

        if row_data:
            client_name = str(row_data.get("nom_client", ""))
            current_statut = str(row_data.get("statut", "Prospect"))

            st.markdown(f"""
            <div class="detail-panel">
                <div class="detail-panel-title">
                    ✏️ Modifier le Deal —
                    <span style="color:{C_CIEL};font-weight:800;">{client_name}</span>
                    &nbsp;·&nbsp; {statut_badge(current_statut)}
                </div>
            </div>
            """, unsafe_allow_html=True)

            # Le form est rendu DANS le container (pas dans le div HTML)
            with st.container():
                with st.form(key=f"edit_deal_{pipeline_id}"):
                    st.markdown(f"""
                    <div style="background:{C_MARINE}08;border:1px solid {C_MARINE}18;
                                border-radius:8px;padding:6px 12px;margin-bottom:12px;
                                font-size:0.8rem;color:{C_MARINE};">
                        Deal ID #{pipeline_id} · Fonds actuel : <b>{row_data.get('fonds','')}</b>
                        · Client : <b>{client_name}</b> ({row_data.get('type_client','')} — {row_data.get('region','')})
                    </div>
                    """, unsafe_allow_html=True)

                    # Ligne 1 : Fonds | Statut
                    r1c1, r1c2 = st.columns(2)
                    with r1c1:
                        fonds_idx = FONDS.index(row_data["fonds"]) if row_data["fonds"] in FONDS else 0
                        new_fonds = st.selectbox("Fonds", FONDS, index=fonds_idx)
                    with r1c2:
                        stat_idx  = STATUTS.index(current_statut) if current_statut in STATUTS else 0
                        new_statut = st.selectbox("Statut", STATUTS, index=stat_idx)

                    # Ligne 2 : AUM
                    r2c1, r2c2, r2c3 = st.columns(3)
                    with r2c1:
                        new_target = st.number_input(
                            "AUM Cible (€)",
                            value=float(row_data.get("target_aum_initial", 0.0)),
                            min_value=0.0, step=1_000_000.0, format="%.0f"
                        )
                    with r2c2:
                        new_revised = st.number_input(
                            "AUM Révisé (€)",
                            value=float(row_data.get("revised_aum", 0.0)),
                            min_value=0.0, step=1_000_000.0, format="%.0f"
                        )
                    with r2c3:
                        new_funded = st.number_input(
                            "AUM Financé (€)",
                            value=float(row_data.get("funded_aum", 0.0)),
                            min_value=0.0, step=1_000_000.0, format="%.0f"
                        )

                    # Ligne 3 : Raison perte | Concurrent | Next Action
                    r3c1, r3c2, r3c3 = st.columns(3)
                    with r3c1:
                        current_raison = str(row_data.get("raison_perte") or "")
                        raison_options = [""] + RAISONS_PERTE
                        raison_idx     = raison_options.index(current_raison) \
                                         if current_raison in raison_options else 0
                        new_raison = st.selectbox(
                            "Raison Perte/Pause" + (" ⚠️ Obligatoire" if new_statut in ("Lost","Paused") else ""),
                            raison_options, index=raison_idx
                        )
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

                    # Bouton submit
                    col_btn, col_info = st.columns([1, 2])
                    with col_btn:
                        submitted = st.form_submit_button(
                            "💾 Sauvegarder les modifications",
                            use_container_width=True
                        )

                if submitted:
                    ok, msg = db.update_pipeline_row({
                        "id":                pipeline_id,
                        "fonds":             new_fonds,
                        "statut":            new_statut,
                        "target_aum_initial": new_target,
                        "revised_aum":        new_revised,
                        "funded_aum":         new_funded,
                        "raison_perte":       new_raison,
                        "concurrent_choisi":  new_concurrent,
                        "next_action_date":   new_nad,
                    })
                    if ok:
                        st.success(
                            f"✅ Deal **{new_fonds}** / **{client_name}** mis à jour "
                            f"→ Statut : **{new_statut}** · AUM Financé : **{fmt_m(new_funded)}**"
                        )
                        st.rerun()
                    else:
                        st.error(msg)
    else:
        st.markdown(f"""
        <div style="background:{C_MARINE}06;border:1px dashed {C_MARINE}30;
                    border-radius:10px;padding:28px;text-align:center;margin-top:14px;">
            <div style="font-size:1.8rem;margin-bottom:8px;">👆</div>
            <div style="color:{C_MARINE};font-weight:600;">
                Sélectionnez un deal dans le tableau ci-dessus
            </div>
            <div style="color:#888;font-size:0.83rem;margin-top:4px;">
                Le formulaire de modification apparaîtra ici
            </div>
        </div>
        """, unsafe_allow_html=True)

    st.divider()

    # ---- Graphique AUM groupé ----
    st.markdown("#### 📊 Comparaison AUM par Deal")

    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    df_viz = df_pipe[
        (df_pipe["target_aum_initial"] > 0) &
        (df_pipe["statut"].isin(["Funded","Soft Commit","Due Diligence","Initial Pitch"]))
    ].copy().head(10)

    if not df_viz.empty:
        fig, ax = plt.subplots(figsize=(12, 4.2))
        fig.patch.set_facecolor("white")
        ax.set_facecolor("white")

        x = np.arange(len(df_viz))
        w = 0.28

        ax.bar(x - w,     df_viz["target_aum_initial"], w,
               label="AUM Cible",   color=C_GRIS,     edgecolor="white")
        ax.bar(x,         df_viz["revised_aum"], w,
               label="AUM Révisé",  color=C_BLEU_MID, edgecolor="white")
        ax.bar(x + w,     df_viz["funded_aum"], w,
               label="AUM Financé", color=C_CIEL,     edgecolor="white")

        ax.set_xticks(x)
        ax.set_xticklabels(df_viz["nom_client"].str[:18],
                           rotation=28, ha="right", fontsize=8.5, color=C_MARINE)
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda y,_: f"${y/1e6:.0f}M"))
        ax.tick_params(axis="y", colors=C_MARINE, labelsize=8)
        ax.set_title("AUM Cible / Révisé / Financé — Deals Actifs",
                     fontsize=10.5, fontweight="bold", color=C_MARINE, pad=8)
        ax.legend(fontsize=8.5, frameon=False, labelcolor=C_MARINE)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.grid(axis="y", alpha=0.25, color=C_GRIS)
        fig.tight_layout()
        st.pyplot(fig, use_container_width=True)
        plt.close(fig)
    else:
        st.info("Pas suffisamment de données actives pour le graphique.")

    # ---- Tableau Lost/Paused ----
    df_lp = df_pipe[df_pipe["statut"].isin(["Lost","Paused"])].copy()
    if not df_lp.empty:
        st.divider()
        st.markdown("#### 🚫 Deals Perdus / En Pause")
        df_lp_disp = df_lp[["nom_client","fonds","statut","target_aum_initial",
                             "raison_perte","concurrent_choisi"]].copy()
        st.dataframe(df_lp_disp, use_container_width=True, hide_index=True,
            column_config={
                "nom_client":         st.column_config.TextColumn("Client"),
                "fonds":              st.column_config.TextColumn("Fonds"),
                "statut":             st.column_config.TextColumn("Statut"),
                "target_aum_initial": st.column_config.NumberColumn("AUM Cible", format="€ %,.0f"),
                "raison_perte":       st.column_config.TextColumn("Raison"),
                "concurrent_choisi":  st.column_config.TextColumn("Concurrent"),
            })


# ============================================================================
# ONGLET 3 : EXECUTIVE DASHBOARD  (inchangé dans sa logique, refactorisé types)
# ============================================================================
with tab_dash:
    st.markdown('<div class="section-title">📊 Executive Dashboard</div>',
                unsafe_allow_html=True)

    kpis = db.get_kpis()

    # KPIs principaux
    kpi_cols = st.columns(4, gap="medium")
    kpi_items = [
        ("AUM Financé Total",  fmt_m(kpis["total_funded"]),   f"{kpis['nb_funded']} deal(s) Funded"),
        ("Pipeline Actif",     fmt_m(kpis["pipeline_actif"]), f"{kpis['nb_deals_actifs']} en cours"),
        ("Taux Conversion",    f"{kpis['taux_conversion']}%", f"{kpis['nb_funded']} funded / {kpis['nb_lost']} lost"),
        ("Deals Actifs",       str(kpis["nb_deals_actifs"]),  "Prospect → Soft Commit"),
    ]
    for col, (lbl, val, sub) in zip(kpi_cols, kpi_items):
        with col:
            st.markdown(kpi_card(lbl, val, sub), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Badges statuts
    STATUT_COLORS = {
        "Funded": C_CIEL, "Soft Commit": C_BLEU_MID, "Due Diligence": "#004F8C",
        "Initial Pitch": "#3A7EBA", "Prospect": "#8FBEDB",
        "Lost": "#AAAAAA", "Paused": "#C0D8E8", "Redeemed": "#B0CEE8",
    }
    statut_order = [s for s in ["Prospect","Initial Pitch","Due Diligence",
                                 "Soft Commit","Funded","Lost","Paused","Redeemed"]
                    if kpis["statut_repartition"].get(s,0) > 0]
    if statut_order:
        badge_cols = st.columns(len(statut_order), gap="small")
        for col, statut in zip(badge_cols, statut_order):
            c_hex = STATUT_COLORS.get(statut, C_CIEL)
            count = kpis["statut_repartition"][statut]
            with col:
                st.markdown(f"""
                <div style="background:{c_hex}22;border:1px solid {c_hex}44;
                            border-radius:8px;padding:10px;text-align:center;">
                    <div style="font-size:0.68rem;color:{C_MARINE};font-weight:700;
                                text-transform:uppercase;">{statut}</div>
                    <div style="font-size:1.55rem;font-weight:800;color:{c_hex};">{count}</div>
                </div>""", unsafe_allow_html=True)

    st.divider()

    # Alertes overdue
    df_overdue = db.get_overdue_actions()
    if not df_overdue.empty:
        st.markdown(f"""
        <div style="background:#FFF8E1;border-left:4px solid {C_CIEL};
                    border-radius:0 8px 8px 0;padding:12px 16px;margin-bottom:16px;">
            <div style="font-size:0.87rem;font-weight:700;color:{C_MARINE};margin-bottom:8px;">
                ⏰ {len(df_overdue)} Action(s) en Retard
            </div>
        """, unsafe_allow_html=True)
        for _, row in df_overdue.iterrows():
            nad = row.get("next_action_date")
            if isinstance(nad, date):
                days_late = (date.today() - nad).days
                nad_str   = nad.isoformat()
            else:
                nad_str   = str(nad or "—")
                days_late = 0
            st.markdown(f"""
            <div class="alert-overdue">
                🔴 <b>{row['nom_client']}</b> — {row['fonds']}
                <span style="color:{C_CIEL};font-weight:600;"> ({row['statut']})</span>
                — Prévue le <b>{nad_str}</b>
                <span style="color:#B04000;"> ({days_late} j. de retard)</span>
            </div>""", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    # Graphiques
    gcol1, gcol2 = st.columns([1, 1.4], gap="large")

    with gcol1:
        st.markdown("#### 🥧 AUM Funded par Type Client")
        if kpis["aum_by_type"]:
            fig_p, ax_p = plt.subplots(figsize=(5, 4.2))
            fig_p.patch.set_facecolor("white")
            lbls = list(kpis["aum_by_type"].keys())
            vals = list(kpis["aum_by_type"].values())
            pal  = [C_CIEL, C_MARINE, C_BLEU_MID, C_BLEU_PALE, C_BLEU_DEEP]
            cols = [pal[i % len(pal)] for i in range(len(lbls))]
            wedges, _, autotexts = ax_p.pie(
                vals, colors=cols, autopct="%1.1f%%", startangle=90,
                pctdistance=0.75,
                wedgeprops={"width":0.55,"edgecolor":"white","linewidth":1.5}
            )
            for at in autotexts:
                at.set_fontsize(8.5); at.set_color("white"); at.set_fontweight("bold")
            tot = sum(vals)
            ax_p.text(0, 0.07, fmt_m(tot), ha="center", va="center",
                      fontsize=11, fontweight="bold", color=C_MARINE)
            ax_p.text(0, -0.18, "Total Funded", ha="center", va="center",
                      fontsize=7.5, color="#666")
            patches = [matplotlib.patches.Patch(color=cols[i],
                                                label=f"{lbls[i]}: {fmt_m(vals[i])}")
                       for i in range(len(lbls))]
            ax_p.legend(handles=patches, loc="lower center", bbox_to_anchor=(0.5,-0.22),
                        ncol=2, fontsize=7.5, frameon=False, labelcolor=C_MARINE)
            fig_p.tight_layout()
            st.pyplot(fig_p, use_container_width=True)
            plt.close(fig_p)
        else:
            st.info("Aucun deal Funded.")

    with gcol2:
        st.markdown("#### 📊 AUM Funded par Fonds")
        if kpis["aum_by_fonds"]:
            fig_b, ax_b = plt.subplots(figsize=(7, 4.2))
            fig_b.patch.set_facecolor("white"); ax_b.set_facecolor("white")
            flbls = list(kpis["aum_by_fonds"].keys())
            fvals = list(kpis["aum_by_fonds"].values())
            yp    = range(len(flbls))
            bars  = ax_b.barh(yp, fvals, color=C_CIEL, edgecolor="white", height=0.55)
            for bar, val in zip(bars, fvals):
                ax_b.text(bar.get_width() + max(fvals)*0.01,
                          bar.get_y() + bar.get_height()/2,
                          fmt_m(val), va="center", ha="left",
                          fontsize=8.5, color=C_MARINE, fontweight="bold")
            ax_b.set_yticks(yp); ax_b.set_yticklabels(flbls, fontsize=9, color=C_MARINE)
            ax_b.invert_yaxis()
            ax_b.xaxis.set_major_formatter(plt.FuncFormatter(lambda x,_: fmt_m(x)))
            ax_b.tick_params(axis="x", colors=C_MARINE, labelsize=8)
            ax_b.set_xlim(0, max(fvals)*1.18)
            ax_b.spines["top"].set_visible(False); ax_b.spines["right"].set_visible(False)
            ax_b.grid(axis="x", alpha=0.25, color=C_GRIS)
            fig_b.tight_layout()
            st.pyplot(fig_b, use_container_width=True)
            plt.close(fig_b)

    st.divider()

    # Top 10 deals
    st.markdown("#### 🏆 Top Deals — AUM Financé")
    df_top = db.get_pipeline_with_clients()
    df_funded = df_top[df_top["statut"] == "Funded"].sort_values("funded_aum", ascending=False).head(10)
    if not df_funded.empty:
        max_f = float(df_funded["funded_aum"].max())
        for i, (_, row) in enumerate(df_funded.iterrows()):
            val = float(row["funded_aum"])
            pct = val / max_f * 100 if max_f > 0 else 0
            medal = ["🥇","🥈","🥉"][i] if i < 3 else f"#{i+1}"
            st.markdown(f"""
            <div style="display:flex;align-items:center;gap:12px;margin:7px 0;
                        padding:10px 14px;background:#F8FAFD;border-radius:8px;
                        border:1px solid {C_MARINE}10;">
                <div style="font-size:1.1rem;min-width:28px;">{medal}</div>
                <div style="flex:1;">
                    <div style="font-size:0.87rem;font-weight:700;color:{C_MARINE};">
                        {row['nom_client']}</div>
                    <div style="font-size:0.73rem;color:#777;">
                        {row['fonds']} · {row['type_client']} · {row['region']}</div>
                    <div style="background:{C_GRIS};border-radius:4px;height:5px;
                                margin-top:5px;overflow:hidden;">
                        <div style="background:{C_CIEL};width:{pct:.0f}%;height:100%;
                                    border-radius:4px;"></div>
                    </div>
                </div>
                <div style="text-align:right;min-width:75px;">
                    <div style="font-size:0.95rem;font-weight:800;color:{C_CIEL};">
                        {fmt_m(val)}</div>
                    <div style="font-size:0.68rem;color:#999;">funded</div>
                </div>
            </div>""", unsafe_allow_html=True)
    else:
        st.info("Aucun deal en statut Funded.")


# ============================================================================
# ONGLET 4 : PERFORMANCE & NAV
# ============================================================================
with tab_perf:
    st.markdown('<div class="section-title">📈 Performance & NAV</div>',
                unsafe_allow_html=True)

    # ---- Instructions & Upload ----
    col_up, col_info = st.columns([1.2, 1], gap="large")

    with col_info:
        st.markdown(f"""
        <div style="background:{C_MARINE}08;border:1px solid {C_MARINE}18;
                    border-radius:10px;padding:18px 20px;">
            <div style="font-weight:700;color:{C_MARINE};font-size:0.95rem;
                        margin-bottom:12px;">📋 Format attendu</div>
            <div style="font-size:0.83rem;color:#444;line-height:1.7;">
                <b>Colonnes obligatoires :</b><br>
                • <code>Date</code> — Format YYYY-MM-DD ou DD/MM/YYYY<br>
                • <code>Fonds</code> — Nom du fonds<br>
                • <code>NAV</code> — Valeur liquidative (numérique)<br><br>
                <b>Calculs générés :</b><br>
                • Courbe Base 100 (normalisée au 1er point)<br>
                • Performance 1 Mois glissant<br>
                • Performance YTD (depuis le 1er janvier)<br>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # Générateur de données de démonstration
        st.markdown("#### 🎲 Données de démonstration")
        if st.button("Générer un fichier NAV de démo", use_container_width=True):
            import io as _io
            demo_dates = pd.date_range(
                start=f"{date.today().year - 1}-01-01",
                end=date.today(),
                freq="B"  # jours ouvrés
            )
            rng = np.random.default_rng(42)
            demo_rows = []
            nav_init = {
                "Global Value":      100.0, "Resilient Equity": 100.0,
                "Income Builder":    100.0, "Private Debt":     100.0,
                "International Fund":100.0,
            }
            for d in demo_dates:
                for fonds, nav in nav_init.items():
                    ret    = rng.normal(0.0003, 0.006)
                    nav   *= (1 + ret)
                    nav_init[fonds] = nav
                    demo_rows.append({"Date": d.date().isoformat(), "Fonds": fonds, "NAV": round(nav, 4)})
            df_demo = pd.DataFrame(demo_rows)
            buf_demo = _io.BytesIO()
            df_demo.to_excel(buf_demo, index=False)
            buf_demo.seek(0)
            st.download_button("⬇️ Télécharger nav_demo.xlsx",
                               data=buf_demo, file_name="nav_demo.xlsx",
                               mime="application/vnd.ms-excel",
                               use_container_width=True)

    with col_up:
        nav_file = st.file_uploader(
            "📂 Charger l'historique NAV (Excel ou CSV)",
            type=["xlsx","xls","csv"],
            help="Colonnes : Date, Fonds, NAV"
        )

    # ---- Traitement du fichier NAV ----
    if nav_file is not None:
        try:
            # Lecture
            df_nav = (pd.read_csv(nav_file)
                      if nav_file.name.endswith(".csv")
                      else pd.read_excel(nav_file))

            # Normalisation des noms de colonnes
            df_nav.columns = [c.strip() for c in df_nav.columns]

            # Vérification colonnes obligatoires
            missing = [c for c in ["Date","Fonds","NAV"] if c not in df_nav.columns]
            if missing:
                st.error(f"❌ Colonnes manquantes : {missing}. "
                         f"Colonnes trouvées : {list(df_nav.columns)}")
                st.stop()

            # Nettoyage strict des types
            df_nav["Date"] = pd.to_datetime(df_nav["Date"], errors="coerce", dayfirst=True)
            df_nav["NAV"]  = pd.to_numeric(df_nav["NAV"], errors="coerce")
            df_nav         = df_nav.dropna(subset=["Date","NAV"])
            df_nav["Fonds"] = df_nav["Fonds"].astype(str).str.strip()
            df_nav         = df_nav.sort_values("Date").reset_index(drop=True)

            fonds_list = sorted(df_nav["Fonds"].unique().tolist())
            date_min   = df_nav["Date"].min()
            date_max   = df_nav["Date"].max()

            st.markdown(f"""
            <div style="background:{C_CIEL}18;border-left:3px solid {C_CIEL};
                        border-radius:0 8px 8px 0;padding:10px 16px;margin:12px 0;">
                ✅ <b>{len(df_nav):,} points</b> chargés · <b>{len(fonds_list)} fonds</b> ·
                Période : <b>{date_min.strftime('%d/%m/%Y')}</b> →
                <b>{date_max.strftime('%d/%m/%Y')}</b>
            </div>
            """, unsafe_allow_html=True)

            # ---- Filtres fonds & période ----
            st.markdown("---")
            ff1, ff2, ff3 = st.columns([2, 1, 1])
            with ff1:
                fonds_sel_nav = st.multiselect(
                    "Fonds à afficher", fonds_list, default=fonds_list[:5]
                )
            with ff2:
                d_debut = st.date_input("Depuis", value=date_min.date())
            with ff3:
                d_fin   = st.date_input("Jusqu'au", value=date_max.date())

            if not fonds_sel_nav:
                st.warning("Sélectionnez au moins un fonds.")
                st.stop()

            # Filtrage
            mask = (
                (df_nav["Fonds"].isin(fonds_sel_nav)) &
                (df_nav["Date"].dt.date >= d_debut) &
                (df_nav["Date"].dt.date <= d_fin)
            )
            df_filtered_nav = df_nav[mask].copy()

            if df_filtered_nav.empty:
                st.warning("Aucune donnée pour la période sélectionnée.")
                st.stop()

            # ---- Calcul Base 100 ----
            pivot = (df_filtered_nav
                     .pivot_table(index="Date", columns="Fonds", values="NAV", aggfunc="last")
                     .sort_index()
                     .ffill())

            # Réindexation sur la période choisie
            pivot = pivot[fonds_sel_nav]

            # Base 100 = normalisé au premier point valide de chaque fonds
            base100 = pivot.copy()
            for fonds in fonds_sel_nav:
                first_valid = pivot[fonds].first_valid_index()
                if first_valid is not None:
                    base100[fonds] = pivot[fonds] / pivot[fonds][first_valid] * 100

            # ---- GRAPHIQUE NAV BASE 100 ----
            st.markdown("#### 📉 Évolution NAV — Base 100")

            fig_nav, ax_nav = plt.subplots(figsize=(12, 5))
            fig_nav.patch.set_facecolor("white")
            ax_nav.set_facecolor("white")

            for i, fonds in enumerate(fonds_sel_nav):
                color    = NAV_PALETTE[i % len(NAV_PALETTE)]
                series   = base100[fonds].dropna()
                line_w   = 2.0 if i < 3 else 1.5
                line_sty = "-" if i % 2 == 0 else "--"

                ax_nav.plot(
                    series.index, series.values,
                    label=fonds,
                    color=color,
                    linewidth=line_w,
                    linestyle=line_sty,
                    alpha=0.92
                )
                # Marqueur sur le dernier point
                if not series.empty:
                    ax_nav.scatter(
                        [series.index[-1]], [series.values[-1]],
                        color=color, s=40, zorder=5
                    )

            # Ligne de référence base 100
            ax_nav.axhline(100, color=C_GRIS, linewidth=1, linestyle=":", alpha=0.7)
            ax_nav.text(
                ax_nav.get_xlim()[0] if ax_nav.get_xlim()[0] != 0 else base100.index[0],
                101, "Base 100", fontsize=7.5, color="#999"
            )

            # Zone colorée au-dessus/dessous de 100
            for i, fonds in enumerate(fonds_sel_nav[:1]):  # zone pour le 1er fonds
                series = base100[fonds].dropna()
                if not series.empty:
                    color = NAV_PALETTE[0]
                    ax_nav.fill_between(
                        series.index, 100, series.values,
                        where=(series.values >= 100),
                        alpha=0.04, color=C_CIEL
                    )
                    ax_nav.fill_between(
                        series.index, 100, series.values,
                        where=(series.values < 100),
                        alpha=0.04, color=C_GRIS
                    )

            ax_nav.set_xlabel("Date", fontsize=9, color=C_MARINE)
            ax_nav.set_ylabel("NAV (Base 100)", fontsize=9, color=C_MARINE)
            ax_nav.tick_params(colors=C_MARINE, labelsize=8)
            ax_nav.spines["top"].set_visible(False)
            ax_nav.spines["right"].set_visible(False)
            ax_nav.spines["left"].set_color(C_GRIS)
            ax_nav.spines["bottom"].set_color(C_GRIS)
            ax_nav.grid(axis="y", alpha=0.25, color=C_GRIS, linewidth=0.7)
            ax_nav.grid(axis="x", alpha=0.1, color=C_GRIS, linewidth=0.5)
            ax_nav.legend(
                fontsize=8.5, frameon=True, framealpha=0.92,
                edgecolor=C_GRIS, labelcolor=C_MARINE,
                loc="upper left"
            )
            ax_nav.set_title(
                f"Performance NAV — Base 100 · {d_debut.strftime('%d/%m/%Y')} → {d_fin.strftime('%d/%m/%Y')}",
                fontsize=11, fontweight="bold", color=C_MARINE, pad=10
            )

            plt.xticks(rotation=20, ha="right")
            fig_nav.tight_layout()
            st.pyplot(fig_nav, use_container_width=True)
            plt.close(fig_nav)

            # ---- TABLEAU PERFORMANCES ----
            st.markdown("#### 📋 Tableau des Performances")

            today_dt   = pd.Timestamp(date.today())
            one_m_ago  = today_dt - pd.DateOffset(months=1)
            jan_1      = pd.Timestamp(f"{date.today().year}-01-01")

            perf_rows = []
            for fonds in fonds_sel_nav:
                series = pivot[fonds].dropna()
                if series.empty:
                    continue

                nav_last = float(series.iloc[-1])
                nav_first = float(series.iloc[0])

                # Perf 1M
                s_1m = series[series.index >= one_m_ago]
                perf_1m = (
                    (nav_last / float(s_1m.iloc[0]) - 1) * 100
                    if len(s_1m) > 0 else float("nan")
                )

                # Perf YTD
                s_ytd = series[series.index >= jan_1]
                perf_ytd = (
                    (nav_last / float(s_ytd.iloc[0]) - 1) * 100
                    if len(s_ytd) > 0 else float("nan")
                )

                # Perf période complète
                perf_period = (nav_last / nav_first - 1) * 100 if nav_first != 0 else float("nan")

                # NAV au dernier point Base100
                nav_b100 = float(base100[fonds].dropna().iloc[-1]) \
                           if not base100[fonds].dropna().empty else float("nan")

                perf_rows.append({
                    "Fonds":            fonds,
                    "NAV Dernière":     round(nav_last, 4),
                    "Base 100 Actuel":  round(nav_b100, 2),
                    "Perf 1M (%)":      round(perf_1m, 2) if not np.isnan(perf_1m) else None,
                    "Perf YTD (%)":     round(perf_ytd, 2) if not np.isnan(perf_ytd) else None,
                    "Perf Période (%)": round(perf_period, 2) if not np.isnan(perf_period) else None,
                })

            if perf_rows:
                df_perf_table = pd.DataFrame(perf_rows)

                # Affichage avec mise en forme conditionnelle HTML
                def perf_fmt(val, suffix="%"):
                    if val is None or (isinstance(val, float) and np.isnan(val)):
                        return '<span style="color:#999;">—</span>'
                    color = C_CIEL if val >= 0 else "#8B2020"
                    sign  = "+" if val > 0 else ""
                    return f'<span style="color:{color};font-weight:700;">{sign}{val:.2f}{suffix}</span>'

                # Rendu HTML du tableau
                table_html = f"""
                <table style="width:100%;border-collapse:collapse;font-size:0.86rem;">
                    <thead>
                        <tr style="background:{C_MARINE};color:{C_BLANC};">
                            <th style="padding:10px 14px;text-align:left;">Fonds</th>
                            <th style="padding:10px 14px;text-align:right;">NAV Dernière</th>
                            <th style="padding:10px 14px;text-align:right;">Base 100</th>
                            <th style="padding:10px 14px;text-align:right;">Perf 1M</th>
                            <th style="padding:10px 14px;text-align:right;">Perf YTD</th>
                            <th style="padding:10px 14px;text-align:right;">Perf Période</th>
                        </tr>
                    </thead>
                    <tbody>
                """
                for i, r in enumerate(perf_rows):
                    bg = "#F8FAFD" if i % 2 == 0 else C_BLANC
                    table_html += f"""
                        <tr style="background:{bg};border-bottom:1px solid {C_GRIS};">
                            <td style="padding:9px 14px;font-weight:600;color:{C_MARINE};">
                                {r['Fonds']}</td>
                            <td style="padding:9px 14px;text-align:right;color:{C_MARINE};">
                                {r['NAV Dernière']:.4f}</td>
                            <td style="padding:9px 14px;text-align:right;color:{C_MARINE};">
                                {r['Base 100 Actuel']:.2f}</td>
                            <td style="padding:9px 14px;text-align:right;">
                                {perf_fmt(r['Perf 1M (%)'])}</td>
                            <td style="padding:9px 14px;text-align:right;">
                                {perf_fmt(r['Perf YTD (%)'])}</td>
                            <td style="padding:9px 14px;text-align:right;">
                                {perf_fmt(r['Perf Période (%)'])}</td>
                        </tr>
                    """
                table_html += "</tbody></table>"
                st.markdown(table_html, unsafe_allow_html=True)

                st.markdown("<br>", unsafe_allow_html=True)

                # Export CSV des perfs
                csv_perf = df_perf_table.to_csv(index=False).encode("utf-8")
                st.download_button(
                    "⬇️ Exporter le tableau en CSV",
                    data=csv_perf,
                    file_name=f"performances_nav_{date.today().isoformat()}.csv",
                    mime="text/csv"
                )

            # ---- Graphique en barres : Perf YTD par fonds ----
            if perf_rows and any(r["Perf YTD (%)"] is not None for r in perf_rows):
                st.markdown("#### 📊 Comparaison Performances YTD")

                ytd_data = [r for r in perf_rows if r["Perf YTD (%)"] is not None]
                ytd_data.sort(key=lambda r: r["Perf YTD (%)"], reverse=True)

                fig_ytd, ax_ytd = plt.subplots(figsize=(10, 3.5))
                fig_ytd.patch.set_facecolor("white")
                ax_ytd.set_facecolor("white")

                fonds_ytd = [r["Fonds"] for r in ytd_data]
                vals_ytd  = [r["Perf YTD (%)"] for r in ytd_data]
                bar_colors = [C_CIEL if v >= 0 else C_GRIS for v in vals_ytd]

                bars_ytd = ax_ytd.bar(
                    range(len(fonds_ytd)), vals_ytd,
                    color=bar_colors, edgecolor="white", linewidth=0.5,
                    width=0.6
                )
                ax_ytd.axhline(0, color=C_MARINE, linewidth=0.8, linestyle="-")

                for bar, val in zip(bars_ytd, vals_ytd):
                    sign = "+" if val > 0 else ""
                    ypos = val + 0.05 if val >= 0 else val - 0.25
                    ax_ytd.text(bar.get_x() + bar.get_width()/2, ypos,
                                f"{sign}{val:.2f}%",
                                ha="center", va="bottom" if val >= 0 else "top",
                                fontsize=8, color=C_MARINE, fontweight="bold")

                ax_ytd.set_xticks(range(len(fonds_ytd)))
                ax_ytd.set_xticklabels(fonds_ytd, rotation=18, ha="right",
                                       fontsize=8.5, color=C_MARINE)
                ax_ytd.set_ylabel("Performance YTD (%)", fontsize=9, color=C_MARINE)
                ax_ytd.tick_params(colors=C_MARINE, labelsize=8)
                ax_ytd.spines["top"].set_visible(False)
                ax_ytd.spines["right"].set_visible(False)
                ax_ytd.grid(axis="y", alpha=0.2, color=C_GRIS)
                ax_ytd.set_title("Performance YTD par Fonds (%)",
                                 fontsize=10, fontweight="bold", color=C_MARINE, pad=8)

                fig_ytd.tight_layout()
                st.pyplot(fig_ytd, use_container_width=True)
                plt.close(fig_ytd)

        except Exception as e:
            st.error(f"❌ Erreur lors du traitement du fichier NAV : {e}")
            import traceback
            with st.expander("🔍 Détails de l'erreur"):
                st.code(traceback.format_exc())

    else:
        # Placeholder quand aucun fichier n'est chargé
        st.markdown(f"""
        <div style="background:{C_MARINE}06;border:2px dashed {C_MARINE}25;
                    border-radius:14px;padding:50px;text-align:center;margin-top:16px;">
            <div style="font-size:2.5rem;margin-bottom:12px;">📈</div>
            <div style="font-size:1.05rem;font-weight:700;color:{C_MARINE};margin-bottom:6px;">
                Module Performance & NAV
            </div>
            <div style="color:#777;font-size:0.86rem;max-width:420px;margin:0 auto;line-height:1.6;">
                Chargez un fichier Excel ou CSV avec les colonnes
                <code>Date</code>, <code>Fonds</code>, <code>NAV</code>
                pour générer les graphiques et le tableau de performances.
                <br><br>
                Utilisez le bouton <b>"Générer un fichier NAV de démo"</b>
                pour obtenir un exemple immédiatement testable.
            </div>
        </div>
        """, unsafe_allow_html=True)
