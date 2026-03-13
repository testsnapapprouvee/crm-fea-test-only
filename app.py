# =============================================================================
# app.py — CRM & Reporting Tool — Asset Management
# Interface Streamlit — Charte Amundi : #002D54 | #00A8E1
# Lancement : streamlit run app.py
# =============================================================================

import streamlit as st
import pandas as pd
import numpy as np
from datetime import date, datetime, timedelta
import io
import sys
import os

# Ajout du répertoire courant au path Python
sys.path.insert(0, os.path.dirname(__file__))

import database as db
import pdf_generator as pdf_gen

# ---------------------------------------------------------------------------
# CONSTANTES CHARTE GRAPHIQUE
# ---------------------------------------------------------------------------
BLEU_MARINE = "#002D54"
BLEU_CIEL   = "#00A8E1"
GRIS_CLAIR  = "#E0E0E0"
BLANC       = "#FFFFFF"

TYPES_CLIENT  = ["IFA", "Wholesale", "Instit", "Family Office"]
REGIONS       = ["GCC", "EMEA", "APAC", "Nordics", "Asia ex-Japan", "North America", "LatAm"]
FONDS         = ["Global Value", "International Fund", "Income Builder",
                 "Resilient Equity", "Private Debt", "Active ETFs"]
STATUTS       = ["Prospect", "Initial Pitch", "Due Diligence",
                 "Soft Commit", "Funded", "Paused", "Lost", "Redeemed"]
STATUTS_ACTIFS = ["Prospect", "Initial Pitch", "Due Diligence", "Soft Commit"]
RAISONS_PERTE  = ["Pricing", "Track Record", "Macro", "Competitor", "Autre"]
TYPES_INTERACTION = ["Call", "Meeting", "Email", "Roadshow", "Conference", "Autre"]


# ---------------------------------------------------------------------------
# CONFIGURATION STREAMLIT
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="CRM Asset Management",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        "Get Help": None,
        "Report a bug": None,
        "About": "CRM & Reporting Tool — Asset Management Division"
    }
)

# ---------------------------------------------------------------------------
# CSS — CHARTE AMUNDI STRICTE
# ---------------------------------------------------------------------------
st.markdown(f"""
<style>
    /* === FOND GLOBAL === */
    .stApp, .main .block-container {{
        background-color: {BLANC};
        color: {BLEU_MARINE};
    }}

    /* === SIDEBAR === */
    [data-testid="stSidebar"] {{
        background-color: {BLEU_MARINE};
    }}
    [data-testid="stSidebar"] * {{
        color: {BLANC} !important;
    }}
    [data-testid="stSidebar"] .stMarkdown h1,
    [data-testid="stSidebar"] .stMarkdown h2,
    [data-testid="stSidebar"] .stMarkdown h3 {{
        color: {BLEU_CIEL} !important;
    }}

    /* === EN-TÊTE PRINCIPAL === */
    .crm-header {{
        background: linear-gradient(135deg, {BLEU_MARINE} 0%, #003D6B 100%);
        color: {BLANC};
        padding: 22px 28px;
        border-radius: 10px;
        margin-bottom: 24px;
        border-left: 5px solid {BLEU_CIEL};
    }}
    .crm-header h1 {{
        color: {BLANC} !important;
        margin: 0;
        font-size: 1.8rem;
        font-weight: 700;
        letter-spacing: -0.3px;
    }}
    .crm-header p {{
        color: {BLEU_CIEL};
        margin: 4px 0 0 0;
        font-size: 0.88rem;
        opacity: 0.9;
    }}

    /* === CARTES KPI === */
    .kpi-card {{
        background: linear-gradient(135deg, {BLEU_MARINE} 0%, #004070 100%);
        color: {BLANC};
        padding: 18px 16px;
        border-radius: 10px;
        text-align: center;
        border: 1px solid {BLEU_CIEL}33;
        transition: transform 0.2s;
    }}
    .kpi-card:hover {{ transform: translateY(-2px); }}
    .kpi-label {{
        font-size: 0.75rem;
        color: {BLEU_CIEL};
        text-transform: uppercase;
        letter-spacing: 0.8px;
        margin-bottom: 8px;
        font-weight: 600;
    }}
    .kpi-value {{
        font-size: 1.75rem;
        font-weight: 800;
        color: {BLANC};
        line-height: 1.1;
    }}
    .kpi-sub {{
        font-size: 0.72rem;
        color: {GRIS_CLAIR};
        margin-top: 4px;
    }}

    /* === ALERTES === */
    .alert-overdue {{
        background-color: #FFF8E1;
        border-left: 4px solid {BLEU_CIEL};
        padding: 10px 14px;
        border-radius: 0 6px 6px 0;
        margin: 6px 0;
        font-size: 0.85rem;
        color: {BLEU_MARINE};
    }}

    /* === SECTION TITLES === */
    .section-title {{
        font-size: 1.1rem;
        font-weight: 700;
        color: {BLEU_MARINE};
        border-bottom: 3px solid {BLEU_CIEL};
        padding-bottom: 6px;
        margin: 20px 0 14px 0;
    }}

    /* === ONGLETS === */
    .stTabs [data-baseweb="tab-list"] {{
        background-color: {BLANC};
        border-bottom: 2px solid {BLEU_MARINE}22;
        gap: 4px;
    }}
    .stTabs [data-baseweb="tab"] {{
        color: {BLEU_MARINE};
        font-weight: 600;
        font-size: 0.88rem;
        padding: 10px 20px;
        border-radius: 6px 6px 0 0;
        background-color: {BLEU_MARINE}0A;
    }}
    .stTabs [aria-selected="true"] {{
        background-color: {BLEU_MARINE} !important;
        color: {BLANC} !important;
    }}

    /* === BOUTONS === */
    .stButton > button {{
        background-color: {BLEU_MARINE};
        color: {BLANC};
        border: none;
        border-radius: 6px;
        font-weight: 600;
        padding: 8px 20px;
        transition: all 0.2s;
    }}
    .stButton > button:hover {{
        background-color: {BLEU_CIEL};
        color: {BLANC};
        transform: translateY(-1px);
        box-shadow: 0 4px 12px {BLEU_CIEL}44;
    }}

    /* === FORMULAIRES === */
    .stSelectbox label, .stTextInput label, .stNumberInput label,
    .stDateInput label, .stTextArea label {{
        color: {BLEU_MARINE} !important;
        font-weight: 600;
        font-size: 0.83rem;
    }}

    /* === DATA EDITOR === */
    [data-testid="stDataEditor"] {{
        border: 1px solid {BLEU_MARINE}22;
        border-radius: 8px;
    }}

    /* === METRIC === */
    [data-testid="stMetric"] {{
        background-color: {BLEU_MARINE}08;
        border: 1px solid {BLEU_MARINE}18;
        border-radius: 8px;
        padding: 12px;
    }}
    [data-testid="stMetricValue"] {{
        color: {BLEU_MARINE} !important;
        font-weight: 800 !important;
    }}
    [data-testid="stMetricLabel"] {{
        color: {BLEU_CIEL} !important;
        font-weight: 600 !important;
    }}

    /* === DIVIDERS === */
    hr {{ border-color: {BLEU_MARINE}22; }}

    /* === SUCCESS / INFO MESSAGES === */
    .stSuccess {{ border-left: 4px solid {BLEU_CIEL}; }}
    .stInfo    {{ border-left: 4px solid {BLEU_MARINE}; }}

    /* === TITRES GÉNÉRAUX === */
    h1, h2, h3, h4 {{ color: {BLEU_MARINE} !important; }}
</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# INITIALISATION BASE DE DONNÉES
# ---------------------------------------------------------------------------
@st.cache_resource
def initialize_database():
    """Initialise la base de données une seule fois."""
    db.init_db()
    return True

initialize_database()


# ---------------------------------------------------------------------------
# HELPERS UI
# ---------------------------------------------------------------------------

def fmt_aum_ui(value: float) -> str:
    """Formate un AUM pour l'interface UI."""
    if value >= 1_000_000_000:
        return f"${value / 1_000_000_000:.2f}Md"
    elif value >= 1_000_000:
        return f"${value / 1_000_000:.1f}M"
    elif value >= 1_000:
        return f"${value / 1_000:.0f}k"
    return f"${value:.0f}"


def color_statut(statut: str) -> str:
    """Retourne une couleur de badge selon le statut (charte Amundi)."""
    mapping = {
        "Funded":        f"background:{BLEU_CIEL};color:{BLANC}",
        "Soft Commit":   f"background:#1A6B9A;color:{BLANC}",
        "Due Diligence": f"background:#004F8C;color:{BLANC}",
        "Initial Pitch": f"background:#3A7EBA;color:{BLANC}",
        "Prospect":      f"background:{BLEU_MARINE}33;color:{BLEU_MARINE}",
        "Lost":          f"background:{GRIS_CLAIR};color:#555",
        "Paused":        f"background:#C8D8E8;color:{BLEU_MARINE}",
        "Redeemed":      f"background:#D0E4F0;color:{BLEU_MARINE}",
    }
    style = mapping.get(statut, f"background:{GRIS_CLAIR};color:#555")
    return f'<span style="padding:3px 10px;border-radius:12px;font-size:0.78rem;font-weight:600;{style}">{statut}</span>'


def kpi_card(label: str, value: str, sub: str = "") -> str:
    """Génère le HTML d'une carte KPI."""
    return f"""
    <div class="kpi-card">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{value}</div>
        {"<div class='kpi-sub'>" + sub + "</div>" if sub else ""}
    </div>
    """


# ---------------------------------------------------------------------------
# SIDEBAR
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown(f"""
    <div style="text-align:center; padding: 10px 0 20px 0;">
        <div style="font-size:2.2rem;">📊</div>
        <div style="font-size:1.1rem; font-weight:800; color:{BLEU_CIEL};">
            CRM Asset Management
        </div>
        <div style="font-size:0.75rem; color:{GRIS_CLAIR}; margin-top:4px;">
            {date.today().strftime("%d %B %Y")}
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.divider()

    # KPIs rapides dans la sidebar
    kpis = db.get_kpis()
    st.markdown(f"""
    <div style="color:{BLEU_CIEL};font-size:0.78rem;font-weight:700;
                text-transform:uppercase;letter-spacing:1px;margin-bottom:8px;">
        📈 Snapshot
    </div>
    """, unsafe_allow_html=True)

    st.markdown(f"""
    <div style="background:{BLANC}15;padding:12px;border-radius:8px;margin-bottom:8px;">
        <div style="font-size:0.72rem;color:{GRIS_CLAIR};">AUM Financé Total</div>
        <div style="font-size:1.3rem;font-weight:800;color:{BLEU_CIEL};">
            {fmt_aum_ui(kpis['total_funded'])}
        </div>
    </div>
    <div style="background:{BLANC}15;padding:12px;border-radius:8px;margin-bottom:8px;">
        <div style="font-size:0.72rem;color:{GRIS_CLAIR};">Pipeline Actif</div>
        <div style="font-size:1.3rem;font-weight:800;color:{BLANC};">
            {fmt_aum_ui(kpis['pipeline_actif'])}
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.divider()

    # --- Export PDF ---
    st.markdown(f"""
    <div style="color:{BLEU_CIEL};font-size:0.78rem;font-weight:700;
                text-transform:uppercase;letter-spacing:1px;margin-bottom:8px;">
        📄 Export Report
    </div>
    """, unsafe_allow_html=True)

    mode_comex = st.toggle(
        "🔒 Mode Comex (Anonymisation)",
        value=False,
        help="Remplace tous les noms clients par {Type} – {Région} dans le PDF et les graphiques."
    )

    if mode_comex:
        st.caption("⚠️ Anonymisation activée : aucun nom client n'apparaîtra dans l'export.")

    if st.button("⬇️ Exporter Executive Report", use_container_width=True):
        with st.spinner("Génération du rapport PDF..."):
            pipeline_df = db.get_pipeline_with_clients()
            fresh_kpis  = db.get_kpis()
            try:
                pdf_bytes = pdf_gen.generate_pdf(
                    pipeline_df=pipeline_df,
                    kpis=fresh_kpis,
                    mode_comex=mode_comex
                )
                filename = (
                    f"executive_report_comex_{date.today().isoformat()}.pdf"
                    if mode_comex
                    else f"executive_report_{date.today().isoformat()}.pdf"
                )
                st.download_button(
                    label="📥 Télécharger le PDF",
                    data=pdf_bytes,
                    file_name=filename,
                    mime="application/pdf",
                    use_container_width=True
                )
                st.success("✅ Rapport généré avec succès.")
            except Exception as e:
                st.error(f"❌ Erreur lors de la génération : {e}")

    st.divider()
    st.caption(f"Version 1.0 — Charte Amundi")


# ---------------------------------------------------------------------------
# EN-TÊTE PRINCIPAL
# ---------------------------------------------------------------------------
st.markdown(f"""
<div class="crm-header">
    <h1>📊 CRM & Reporting — Asset Management</h1>
    <p>Gestion du pipeline commercial | Suivi des mandats | Reporting Exécutif</p>
</div>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# ONGLETS PRINCIPAUX
# ---------------------------------------------------------------------------
tab_ingestion, tab_pipeline, tab_dashboard = st.tabs([
    "📥  Data Ingestion",
    "🔄  Pipeline Management",
    "📊  Executive Dashboard"
])


# ============================================================================
# ONGLET 1 : DATA INGESTION
# ============================================================================
with tab_ingestion:
    st.markdown('<div class="section-title">📥 Saisie & Import de Données</div>',
                unsafe_allow_html=True)

    col_form, col_import = st.columns([1, 1], gap="large")

    # --- Formulaire manuel Clients ---
    with col_form:
        with st.expander("➕ Ajouter un Client", expanded=True):
            with st.form("form_add_client", clear_on_submit=True):
                st.markdown("**Nouveau Client**")
                col1, col2 = st.columns(2)
                with col1:
                    nom_client  = st.text_input("Nom du Client *", placeholder="Ex: Deutsche AM Paris")
                    type_client = st.selectbox("Type Client *", TYPES_CLIENT)
                with col2:
                    region = st.selectbox("Région *", REGIONS)

                submitted_client = st.form_submit_button(
                    "✅ Enregistrer le Client",
                    use_container_width=True
                )

            if submitted_client:
                if not nom_client.strip():
                    st.error("⚠️ Le nom du client est obligatoire.")
                else:
                    try:
                        db.add_client(nom_client.strip(), type_client, region)
                        st.success(f"✅ Client **{nom_client}** ajouté avec succès.")
                        st.cache_resource.clear()
                    except Exception as e:
                        if "UNIQUE" in str(e):
                            st.warning(f"⚠️ Un client nommé **{nom_client}** existe déjà.")
                        else:
                            st.error(f"❌ Erreur : {e}")

        # --- Formulaire manuel Pipeline ---
        with st.expander("➕ Ajouter un Deal Pipeline", expanded=True):
            clients_dict = db.get_client_options()
            if not clients_dict:
                st.info("Ajoutez d'abord un client pour créer un deal.")
            else:
                with st.form("form_add_deal", clear_on_submit=True):
                    st.markdown("**Nouveau Deal**")

                    col_a, col_b = st.columns(2)
                    with col_a:
                        client_sel   = st.selectbox("Client *", list(clients_dict.keys()))
                        fonds_sel    = st.selectbox("Fonds *", FONDS)
                        statut_sel   = st.selectbox("Statut *", STATUTS)
                    with col_b:
                        target_aum   = st.number_input("Target AUM Initial (€)", min_value=0, step=1_000_000)
                        revised_aum  = st.number_input("AUM Révisé (€)", min_value=0, step=1_000_000)
                        funded_aum   = st.number_input("AUM Financé (€)", min_value=0, step=1_000_000)

                    col_c, col_d = st.columns(2)
                    with col_c:
                        raison_perte = ""
                        concurrent   = ""
                        if statut_sel in ("Lost", "Paused"):
                            raison_perte = st.selectbox("Raison Perte/Pause *", RAISONS_PERTE)
                            concurrent   = st.text_input("Concurrent Choisi")
                        else:
                            st.caption("*(Raison perte : non applicable)*")
                    with col_d:
                        next_action  = st.date_input(
                            "Prochaine Action",
                            value=date.today() + timedelta(days=14)
                        )

                    sub_deal = st.form_submit_button("✅ Enregistrer le Deal", use_container_width=True)

                if sub_deal:
                    if statut_sel in ("Lost", "Paused") and not raison_perte:
                        st.error("⚠️ La raison de perte/pause est obligatoire.")
                    else:
                        client_id = clients_dict[client_sel]
                        try:
                            db.add_pipeline_entry(
                                client_id, fonds_sel, statut_sel,
                                float(target_aum), float(revised_aum), float(funded_aum),
                                raison_perte, concurrent, next_action.isoformat()
                            )
                            st.success(f"✅ Deal **{fonds_sel}** pour **{client_sel}** ajouté.")
                        except Exception as e:
                            st.error(f"❌ Erreur : {e}")

        # --- Formulaire Activités ---
        with st.expander("📝 Enregistrer une Activité"):
            clients_dict2 = db.get_client_options()
            if clients_dict2:
                with st.form("form_add_activity", clear_on_submit=True):
                    col_act1, col_act2 = st.columns(2)
                    with col_act1:
                        act_client = st.selectbox("Client *", list(clients_dict2.keys()))
                        act_type   = st.selectbox("Type d'Interaction", TYPES_INTERACTION)
                    with col_act2:
                        act_date   = st.date_input("Date", value=date.today())
                        act_notes  = st.text_area("Notes / Compte-rendu", height=80)

                    sub_act = st.form_submit_button("✅ Enregistrer l'Activité", use_container_width=True)

                if sub_act:
                    db.add_activity(clients_dict2[act_client], act_date.isoformat(), act_notes, act_type)
                    st.success(f"✅ Activité enregistrée pour **{act_client}**.")

    # --- Import CSV/Excel ---
    with col_import:
        st.markdown("#### 📂 Import CSV / Excel (Upsert)")
        st.caption("La logique Upsert met à jour les enregistrements existants et crée les nouveaux.")

        import_type = st.radio(
            "Table cible",
            ["Clients", "Pipeline"],
            horizontal=True
        )

        uploaded_file = st.file_uploader(
            "Glissez un fichier CSV ou Excel (.xlsx)",
            type=["csv", "xlsx", "xls"],
            help="Les colonnes doivent correspondre aux champs de la base."
        )

        if upload_type := import_type:
            if upload_type == "Clients":
                st.info("📋 **Colonnes attendues** : `nom_client`, `type_client`, `region`")
            else:
                st.info("📋 **Colonnes attendues** : `nom_client`, `fonds`, `statut`, "
                        "`target_aum_initial`, `revised_aum`, `funded_aum`, "
                        "`raison_perte`, `concurrent_choisi`, `next_action_date`")

        if uploaded_file is not None:
            try:
                if uploaded_file.name.endswith(".csv"):
                    df_import = pd.read_csv(uploaded_file)
                else:
                    df_import = pd.read_excel(uploaded_file)

                st.markdown("**Aperçu des données :**")
                st.dataframe(
                    df_import.head(5),
                    use_container_width=True,
                    height=180
                )
                st.caption(f"{len(df_import)} ligne(s) détectée(s)")

                if st.button("🔄 Lancer l'Import (Upsert)", use_container_width=True):
                    with st.spinner("Import en cours..."):
                        if import_type == "Clients":
                            ins, upd = db.upsert_clients_from_df(df_import)
                        else:
                            ins, upd = db.upsert_pipeline_from_df(df_import)

                    st.success(
                        f"✅ Import terminé : **{ins} créé(s)**, **{upd} mis à jour**."
                    )

            except Exception as e:
                st.error(f"❌ Erreur lors de la lecture du fichier : {e}")

        st.divider()

        # --- Aperçu activités récentes ---
        st.markdown("#### 📅 Dernières Activités")
        df_act = db.get_activities()
        if not df_act.empty:
            st.dataframe(
                df_act[["nom_client", "date", "type_interaction", "notes"]].head(10),
                use_container_width=True,
                height=280,
                hide_index=True,
                column_config={
                    "nom_client":       st.column_config.TextColumn("Client"),
                    "date":             st.column_config.DateColumn("Date"),
                    "type_interaction": st.column_config.TextColumn("Type"),
                    "notes":            st.column_config.TextColumn("Notes"),
                }
            )
        else:
            st.info("Aucune activité enregistrée.")


# ============================================================================
# ONGLET 2 : PIPELINE MANAGEMENT
# ============================================================================
with tab_pipeline:
    st.markdown('<div class="section-title">🔄 Gestion du Pipeline Commercial</div>',
                unsafe_allow_html=True)

    # --- Filtres ---
    with st.expander("🔍 Filtres", expanded=False):
        filt_col1, filt_col2, filt_col3 = st.columns(3)
        with filt_col1:
            filt_statuts = st.multiselect(
                "Statuts",
                STATUTS,
                default=STATUTS_ACTIFS,
                help="Laisser vide pour tout afficher"
            )
        with filt_col2:
            filt_fonds = st.multiselect("Fonds", FONDS)
        with filt_col3:
            filt_regions = st.multiselect("Régions", REGIONS)

    # Chargement du pipeline
    df_pipeline = db.get_pipeline_with_clients()

    # Application des filtres
    df_filtered = df_pipeline.copy()
    if filt_statuts:
        df_filtered = df_filtered[df_filtered["statut"].isin(filt_statuts)]
    if filt_fonds:
        df_filtered = df_filtered[df_filtered["fonds"].isin(filt_fonds)]
    if filt_regions:
        df_filtered = df_filtered[df_filtered["region"].isin(filt_regions)]

    st.caption(f"**{len(df_filtered)} deal(s)** affiché(s) sur {len(df_pipeline)} total")

    # --- Data Editor ---
    st.markdown("#### ✏️ Éditeur de Pipeline")
    st.info(
        "💡 Modifiez directement les cellules puis cliquez sur **Sauvegarder les Modifications**. "
        "⚠️ Si vous passez un statut en **Lost** ou **Paused**, la colonne *Raison Perte* est **obligatoire**."
    )

    # Colonnes éditables (on masque les colonnes de jointure)
    cols_display = [
        "id", "nom_client", "type_client", "region", "fonds", "statut",
        "target_aum_initial", "revised_aum", "funded_aum",
        "raison_perte", "concurrent_choisi", "next_action_date"
    ]
    df_edit_source = df_filtered[cols_display].copy()

    edited_df = st.data_editor(
        df_edit_source,
        use_container_width=True,
        height=400,
        hide_index=True,
        num_rows="fixed",
        column_config={
            "id": st.column_config.NumberColumn("ID", disabled=True, width="small"),
            "nom_client": st.column_config.TextColumn(
                "Client", disabled=True, width="medium"
            ),
            "type_client": st.column_config.TextColumn(
                "Type", disabled=True, width="small"
            ),
            "region": st.column_config.TextColumn(
                "Région", disabled=True, width="small"
            ),
            "fonds": st.column_config.SelectboxColumn(
                "Fonds", options=FONDS, width="medium"
            ),
            "statut": st.column_config.SelectboxColumn(
                "Statut", options=STATUTS, width="medium"
            ),
            "target_aum_initial": st.column_config.NumberColumn(
                "AUM Cible (€)", format="€ %,.0f", min_value=0, step=1_000_000
            ),
            "revised_aum": st.column_config.NumberColumn(
                "AUM Révisé (€)", format="€ %,.0f", min_value=0, step=1_000_000
            ),
            "funded_aum": st.column_config.NumberColumn(
                "AUM Financé (€)", format="€ %,.0f", min_value=0, step=1_000_000
            ),
            "raison_perte": st.column_config.SelectboxColumn(
                "Raison Perte",
                options=[""] + RAISONS_PERTE,
                help="Obligatoire si statut = Lost ou Paused"
            ),
            "concurrent_choisi": st.column_config.TextColumn("Concurrent"),
            "next_action_date": st.column_config.DateColumn(
                "Prochaine Action", format="DD/MM/YYYY"
            ),
        },
        key="pipeline_editor"
    )

    if st.button("💾 Sauvegarder les Modifications", use_container_width=False):
        with st.spinner("Sauvegarde en cours..."):
            errors = db.bulk_update_pipeline(edited_df)
        if errors:
            for err in errors:
                st.error(err)
        else:
            st.success(f"✅ {len(edited_df)} deal(s) mis à jour avec succès.")

    st.divider()

    # --- Visualisation AUM : Target vs Funded ---
    st.markdown("#### 📊 Historisation Visuelle — AUM Target vs Funded")

    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    df_viz = df_pipeline[
        (df_pipeline["target_aum_initial"] > 0) &
        (df_pipeline["statut"].isin(["Funded", "Soft Commit", "Due Diligence", "Initial Pitch"]))
    ].copy().head(10)

    if not df_viz.empty:
        fig, ax = plt.subplots(figsize=(12, 4.5))
        fig.patch.set_facecolor("white")
        ax.set_facecolor("white")

        x = np.arange(len(df_viz))
        w = 0.32

        bars_target = ax.bar(
            x - w/2, df_viz["target_aum_initial"], w,
            label="Target AUM Initial", color=GRIS_CLAIR, edgecolor="white"
        )
        bars_revised = ax.bar(
            x + w/2, df_viz["revised_aum"], w,
            label="AUM Révisé", color="#5BA3C9", edgecolor="white"
        )
        bars_funded = ax.bar(
            x + w*1.5, df_viz["funded_aum"], w,
            label="AUM Financé", color=BLEU_CIEL, edgecolor="white"
        )

        ax.set_xticks(x + w/2)
        ax.set_xticklabels(
            df_viz["nom_client"].str[:18],
            rotation=30, ha="right", fontsize=8.5, color=BLEU_MARINE
        )
        ax.yaxis.set_major_formatter(
            plt.FuncFormatter(lambda y, _: f"${y/1e6:.0f}M")
        )
        ax.tick_params(axis="y", colors=BLEU_MARINE, labelsize=8)
        ax.set_title("Comparaison AUM Target / Révisé / Financé",
                     fontsize=11, fontweight="bold", color=BLEU_MARINE, pad=10)
        ax.set_ylabel("AUM (€)", fontsize=9, color=BLEU_MARINE)
        ax.legend(fontsize=8.5, frameon=False, labelcolor=BLEU_MARINE)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.grid(axis="y", alpha=0.3, color=GRIS_CLAIR)

        fig.tight_layout()
        st.pyplot(fig, use_container_width=True)
        plt.close(fig)
    else:
        st.info("Pas assez de données pour afficher le graphique.")

    # --- Tableau récap Lost / Paused ---
    df_lp = df_pipeline[df_pipeline["statut"].isin(["Lost", "Paused"])].copy()
    if not df_lp.empty:
        st.markdown("#### 🚫 Deals Perdus / En Pause")
        st.dataframe(
            df_lp[["nom_client", "fonds", "statut", "target_aum_initial",
                   "raison_perte", "concurrent_choisi"]],
            use_container_width=True,
            hide_index=True,
            column_config={
                "nom_client":          st.column_config.TextColumn("Client"),
                "fonds":               st.column_config.TextColumn("Fonds"),
                "statut":              st.column_config.TextColumn("Statut"),
                "target_aum_initial":  st.column_config.NumberColumn("AUM Cible (€)", format="€ %,.0f"),
                "raison_perte":        st.column_config.TextColumn("Raison"),
                "concurrent_choisi":   st.column_config.TextColumn("Concurrent"),
            }
        )


# ============================================================================
# ONGLET 3 : EXECUTIVE DASHBOARD
# ============================================================================
with tab_dashboard:
    st.markdown('<div class="section-title">📊 Executive Dashboard</div>',
                unsafe_allow_html=True)

    # Rafraîchissement des KPIs
    kpis = db.get_kpis()

    # --- LIGNE 1 : KPIs Principaux ---
    kpi_cols = st.columns(4, gap="medium")
    kpi_data_main = [
        ("AUM Financé Total",  fmt_aum_ui(kpis["total_funded"]),
         f"{kpis['nb_funded']} deal(s) Funded"),
        ("Pipeline Actif",     fmt_aum_ui(kpis["pipeline_actif"]),
         f"{kpis['nb_deals_actifs']} deal(s) en cours"),
        ("Taux de Conversion", f"{kpis['taux_conversion']}%",
         f"{kpis['nb_funded']} funded / {kpis['nb_lost']} lost"),
        ("Deals Actifs",       str(kpis["nb_deals_actifs"]),
         "Prospect → Soft Commit"),
    ]
    for col, (label, value, sub) in zip(kpi_cols, kpi_data_main):
        with col:
            st.markdown(kpi_card(label, value, sub), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # --- LIGNE 2 : KPIs secondaires ---
    sec_cols = st.columns(len(kpis["statut_repartition"]), gap="small")
    statut_order = ["Prospect", "Initial Pitch", "Due Diligence",
                    "Soft Commit", "Funded", "Lost", "Paused", "Redeemed"]
    statut_colors = {
        "Funded":        BLEU_CIEL,
        "Soft Commit":   "#1A6B9A",
        "Due Diligence": "#004F8C",
        "Initial Pitch": "#3A7EBA",
        "Prospect":      "#8FBEDB",
        "Lost":          "#AAAAAA",
        "Paused":        "#C0D8E8",
        "Redeemed":      "#B0CEE8",
    }
    for i, statut in enumerate(statut_order):
        count = kpis["statut_repartition"].get(statut, 0)
        if count > 0 and i < len(sec_cols):
            with sec_cols[min(i, len(sec_cols)-1)]:
                color = statut_colors.get(statut, BLEU_CIEL)
                st.markdown(f"""
                <div style="background:{color}22;border:1px solid {color}44;
                            border-radius:8px;padding:10px;text-align:center;">
                    <div style="font-size:0.7rem;color:{BLEU_MARINE};
                                font-weight:700;text-transform:uppercase;">{statut}</div>
                    <div style="font-size:1.6rem;font-weight:800;color:{color};">{count}</div>
                </div>
                """, unsafe_allow_html=True)

    st.divider()

    # --- ALERTES : Next Action Date dépassées ---
    df_overdue = db.get_overdue_actions()
    if not df_overdue.empty:
        st.markdown(f"""
        <div style="background:#FFF8E1;border-left:4px solid {BLEU_CIEL};
                    border-radius:0 8px 8px 0;padding:12px 16px;margin-bottom:16px;">
            <div style="font-size:0.88rem;font-weight:700;color:{BLEU_MARINE};margin-bottom:8px;">
                ⏰ {len(df_overdue)} Action(s) en Retard — Intervention Requise
            </div>
        """, unsafe_allow_html=True)

        for _, row in df_overdue.iterrows():
            days_overdue = (date.today() - date.fromisoformat(str(row["next_action_date"]))).days
            st.markdown(f"""
            <div class="alert-overdue">
                🔴 <b>{row['nom_client']}</b> — {row['fonds']}
                <span style="color:{BLEU_CIEL};font-weight:600;"> ({row['statut']})</span>
                — Action prévue le <b>{row['next_action_date']}</b>
                <span style="color:#B04000;"> ({days_overdue} jour(s) de retard)</span>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("</div>", unsafe_allow_html=True)

    # --- GRAPHIQUES DASHBOARD ---
    chart_col1, chart_col2 = st.columns([1, 1.4], gap="large")

    # Pie Chart — AUM par Type de Client
    with chart_col1:
        st.markdown("#### 🥧 AUM Funded par Type Client")
        if kpis["aum_by_type"]:
            fig_pie, ax_pie = plt.subplots(figsize=(5, 4.2))
            fig_pie.patch.set_facecolor("white")

            labels = list(kpis["aum_by_type"].keys())
            values = list(kpis["aum_by_type"].values())
            palette = [BLEU_CIEL, BLEU_MARINE, "#1A6B9A", "#7BC8E8", "#003F7A"]
            colors  = [palette[i % len(palette)] for i in range(len(labels))]

            wedges, texts, autotexts = ax_pie.pie(
                values,
                labels=None,
                colors=colors,
                autopct="%1.1f%%",
                startangle=90,
                pctdistance=0.75,
                wedgeprops={"width": 0.55, "edgecolor": "white", "linewidth": 1.5},
            )
            for at in autotexts:
                at.set_fontsize(8.5)
                at.set_color("white")
                at.set_fontweight("bold")

            total = sum(values)
            ax_pie.text(0, 0.07, fmt_aum_ui(total), ha="center", va="center",
                        fontsize=11, fontweight="bold", color=BLEU_MARINE)
            ax_pie.text(0, -0.18, "Total Funded", ha="center", va="center",
                        fontsize=7.5, color="#666")

            legend_patches = [
                matplotlib.patches.Patch(color=colors[i],
                                         label=f"{labels[i]}: {fmt_aum_ui(values[i])}")
                for i in range(len(labels))
            ]
            ax_pie.legend(handles=legend_patches, loc="lower center",
                          bbox_to_anchor=(0.5, -0.22), ncol=2, fontsize=7.5,
                          frameon=False, labelcolor=BLEU_MARINE)
            fig_pie.tight_layout()
            st.pyplot(fig_pie, use_container_width=True)
            plt.close(fig_pie)
        else:
            st.info("Aucun deal Funded.")

    # Bar Chart — AUM par Fonds
    with chart_col2:
        st.markdown("#### 📊 AUM Funded par Fonds")
        if kpis["aum_by_fonds"]:
            fig_bar, ax_bar = plt.subplots(figsize=(7, 4.2))
            fig_bar.patch.set_facecolor("white")
            ax_bar.set_facecolor("white")

            fonds_labels = list(kpis["aum_by_fonds"].keys())
            fonds_values = list(kpis["aum_by_fonds"].values())

            y_pos = range(len(fonds_labels))
            bars = ax_bar.barh(
                y_pos, fonds_values,
                color=BLEU_CIEL, edgecolor="white", height=0.55
            )

            for bar, val in zip(bars, fonds_values):
                ax_bar.text(
                    bar.get_width() + max(fonds_values) * 0.01,
                    bar.get_y() + bar.get_height() / 2,
                    fmt_aum_ui(val),
                    va="center", ha="left",
                    fontsize=8.5, color=BLEU_MARINE, fontweight="bold"
                )

            ax_bar.set_yticks(y_pos)
            ax_bar.set_yticklabels(fonds_labels, fontsize=9, color=BLEU_MARINE)
            ax_bar.invert_yaxis()
            ax_bar.set_xlabel("AUM Financé (€)", fontsize=9, color=BLEU_MARINE)
            ax_bar.xaxis.set_major_formatter(
                plt.FuncFormatter(lambda x, _: fmt_aum_ui(x))
            )
            ax_bar.tick_params(axis="x", colors=BLEU_MARINE, labelsize=8)
            ax_bar.set_xlim(0, max(fonds_values) * 1.18)
            ax_bar.spines["top"].set_visible(False)
            ax_bar.spines["right"].set_visible(False)
            ax_bar.grid(axis="x", alpha=0.3, color=GRIS_CLAIR)

            for i, bar in enumerate(bars):
                ax_bar.barh(
                    bar.get_y() + bar.get_height() / 2,
                    max(fonds_values) * 1.15,
                    height=bar.get_height(),
                    color=GRIS_CLAIR if i % 2 == 0 else "white",
                    alpha=0.25, zorder=0
                )

            fig_bar.tight_layout()
            st.pyplot(fig_bar, use_container_width=True)
            plt.close(fig_bar)
        else:
            st.info("Aucun AUM Funded enregistré.")

    st.divider()

    # --- TOP 10 DEALS ---
    st.markdown("#### 🏆 Top Deals — AUM Financé")

    df_top = db.get_pipeline_with_clients()
    df_top_funded = df_top[df_top["statut"] == "Funded"].sort_values(
        "funded_aum", ascending=False
    ).head(10)

    if not df_top_funded.empty:
        max_funded = df_top_funded["funded_aum"].max()

        for i, (_, row) in enumerate(df_top_funded.iterrows()):
            pct = (row["funded_aum"] / max_funded * 100) if max_funded > 0 else 0
            medal = ["🥇", "🥈", "🥉"][i] if i < 3 else f"#{i+1}"

            st.markdown(f"""
            <div style="display:flex;align-items:center;gap:12px;margin:8px 0;
                        padding:10px 14px;background:#F8FAFD;border-radius:8px;
                        border:1px solid {BLEU_MARINE}12;">
                <div style="font-size:1.2rem;min-width:30px;">{medal}</div>
                <div style="flex:1;">
                    <div style="font-size:0.88rem;font-weight:700;color:{BLEU_MARINE};">
                        {row['nom_client']}
                    </div>
                    <div style="font-size:0.75rem;color:#777;">
                        {row['fonds']} · {row['type_client']} · {row['region']}
                    </div>
                    <div style="background:{GRIS_CLAIR};border-radius:4px;height:6px;
                                margin-top:5px;overflow:hidden;">
                        <div style="background:{BLEU_CIEL};width:{pct:.0f}%;height:100%;
                                    border-radius:4px;"></div>
                    </div>
                </div>
                <div style="text-align:right;min-width:80px;">
                    <div style="font-size:1rem;font-weight:800;color:{BLEU_CIEL};">
                        {fmt_aum_ui(row['funded_aum'])}
                    </div>
                    <div style="font-size:0.7rem;color:#999;">funded</div>
                </div>
            </div>
            """, unsafe_allow_html=True)
    else:
        st.info("Aucun deal en statut Funded.")

    st.divider()

    # --- PIPELINE ACTIF COMPLET ---
    st.markdown("#### 📋 Pipeline Actif — Vue Complète")
    df_actif = df_pipeline[df_pipeline["statut"].isin(STATUTS_ACTIFS)].copy()

    if not df_actif.empty:
        # Colonne alerte visuelle
        today = date.today()
        def fmt_next_action(val):
            if pd.isna(val) or not val:
                return "—"
            try:
                d = date.fromisoformat(str(val))
                if d < today:
                    return f"⚠️ {val}"
                elif d <= today + timedelta(days=7):
                    return f"🟡 {val}"
                return str(val)
            except Exception:
                return str(val)

        df_actif["next_action_display"] = df_actif["next_action_date"].apply(fmt_next_action)

        st.dataframe(
            df_actif[["nom_client", "type_client", "region", "fonds", "statut",
                       "revised_aum", "funded_aum", "next_action_display"]],
            use_container_width=True,
            height=350,
            hide_index=True,
            column_config={
                "nom_client":          st.column_config.TextColumn("Client"),
                "type_client":         st.column_config.TextColumn("Type"),
                "region":              st.column_config.TextColumn("Région"),
                "fonds":               st.column_config.TextColumn("Fonds"),
                "statut":              st.column_config.TextColumn("Statut"),
                "revised_aum":         st.column_config.NumberColumn("AUM Révisé", format="€ %,.0f"),
                "funded_aum":          st.column_config.NumberColumn("AUM Financé", format="€ %,.0f"),
                "next_action_display": st.column_config.TextColumn("Prochaine Action"),
            }
        )
    else:
        st.info("Le pipeline actif est vide.")
