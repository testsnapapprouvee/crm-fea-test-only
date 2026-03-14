# =============================================================================
# app.py  —  CRM Asset Management  —  Amundi Edition
# Priorite : STABILITE ABSOLUE
# Charte : #001c4b Marine | #019ee1 Ciel | #f07d00 Orange
# Regles de stabilite :
#   - Plotly EXPLICITE : zero **helper_dict, zero full_figure_for_development
#   - Chaque fig.update_layout() ecrit a plat avec tous les params
#   - st.dialog pour les modales KPI drill-down
#   - Bouton "Backup Excel" en haut de sidebar
#   - Matplotlib ABSENT de ce fichier (reserve a pdf_generator.py)
# Lancement : streamlit run app.py
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
# COULEURS — variables directes, pas de dict helper
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
STATUTS           = ["Prospect", "Initial Pitch", "Due Diligence",
                     "Soft Commit", "Funded", "Paused", "Lost", "Redeemed"]
STATUTS_ACTIFS    = ["Prospect", "Initial Pitch", "Due Diligence", "Soft Commit"]
RAISONS_PERTE     = ["Pricing", "Track Record", "Macro", "Competitor", "Autre"]
TYPES_INTERACTION = ["Call", "Meeting", "Email", "Roadshow", "Conference", "Autre"]

STATUT_COLORS = {
    "Funded":        B_MID,   "Soft Commit":   "#2c7fb8",
    "Due Diligence": "#004f8c", "Initial Pitch": B_PAL,
    "Prospect":      "#9ecae1", "Lost":          "#aaaaaa",
    "Paused":        "#c0c0c0", "Redeemed":      "#b8b8d0",
}


# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------

st.set_page_config(page_title="CRM - Asset Management",
                   layout="wide", initial_sidebar_state="expanded")


# ---------------------------------------------------------------------------
# CSS — charte stricte, code plat
# ---------------------------------------------------------------------------

st.markdown("""
<style>
/* =====================================================================
   CHARTE AMUNDI — Marine #001c4b | Ciel #019ee1 | Orange #f07d00 (unique)
   Principe : zero border-radius, design institutionnel angulaire
   ===================================================================== */

.stApp, .main .block-container {
    background-color: #ffffff;
    color: #001c4b;
    font-family: 'Segoe UI', Arial, sans-serif;
}

/* Sidebar — fond marine plein, aucun arrondi */
[data-testid="stSidebar"] {
    background-color: #001c4b;
}
[data-testid="stSidebar"] * { color: #ffffff !important; }
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 { color: #019ee1 !important; }
[data-testid="stSidebar"] .stMultiSelect > div,
[data-testid="stSidebar"] .stSelectbox > div { !important; }

/* En-tete application */
.crm-header {
    background: #001c4b;
    padding: 16px 24px;
    margin-bottom: 16px;
    border-bottom: 3px solid #019ee1;
}
.crm-header h1 {
    color: #ffffff !important;
    margin: 0;
    font-size: 1.4rem;
    font-weight: 700;
}
.crm-header p { color: #7ab8d8; margin: 3px 0 0 0; font-size: 0.80rem; }

/* KPI Cards — fond marine, aucun arrondi, cliquable via JS */
.kpi-card {
    background: #001c4b;
    padding: 13px 10px;
    text-align: center;
    border: 1px solid #1a5e8a;
    cursor: pointer;
}
.kpi-card:hover { background: #0a2d5e; }
.kpi-label {
    font-size: 0.64rem;
    color: #7ab8d8;
    text-transform: uppercase;
    letter-spacing: 0.8px;
    margin-bottom: 5px;
    font-weight: 600;
}
.kpi-value { font-size: 1.4rem; font-weight: 800; color: #ffffff; }
.kpi-sub   { font-size: 0.61rem; color: #c8dde8; margin-top: 3px; }

/* Titres de sections */
.section-title {
    font-size: 0.90rem;
    font-weight: 700;
    color: #001c4b;
    border-bottom: 2px solid #001c4b22;
    padding-bottom: 4px;
    margin: 14px 0 9px 0;
}

/* Badge RETARD — orange, seul usage orange dans l'UI */
.badge-retard {
    display: inline-block;
    background: #f07d00;
    color: #ffffff;
    padding: 1px 7px;
    font-size: 0.66rem;
    font-weight: 700;
    letter-spacing: 0.4px;
}

/* Badge perimetre fonds — sidebar */
.perimetre-badge {
    display: inline-block;
    background: #019ee122;
    border: 1px solid #019ee155;
    padding: 2px 8px;
    font-size: 0.70rem;
    color: #ffffff;
    font-weight: 600;
    margin: 2px;
}

/* Panneau de detail pipeline */
.detail-panel {
    background: #f2f6fa;
    border-left: 3px solid #019ee1;
    padding: 14px 16px 11px 16px;
    margin-top: 10px;
}

/* Alertes retard */
.alert-overdue {
    background: #fef6f0;
    border-left: 3px solid #f07d00;
    padding: 6px 11px;
    margin: 3px 0;
    font-size: 0.77rem;
    color: #001c4b;
}

/* Card commercial */
.sales-card {
    background: #f4f8fc;
    border: 1px solid #001c4b18;
    padding: 13px;
    border-top: 3px solid #019ee1;
}
.sales-card-name {
    font-size: 0.87rem;
    font-weight: 700;
    color: #001c4b;
    margin-bottom: 8px;
    padding-bottom: 5px;
    border-bottom: 1px solid #e8e8e8;
}
.sales-metric { font-size: 0.69rem; color: #666; margin-bottom: 2px; }
.sales-metric-val { font-size: 0.93rem; font-weight: 700; color: #001c4b; }
.sales-metric-acc { font-size: 0.93rem; font-weight: 700; color: #019ee1; }

/* Onglets — angulaires */
.stTabs [data-baseweb="tab-list"] {
    background: #f0f4f8;
    border-bottom: 2px solid #001c4b20;
    gap: 0;
}
.stTabs [data-baseweb="tab"] {
    color: #001c4b;
    font-weight: 600;
    font-size: 0.81rem;
    padding: 7px 16px;
    background: #f0f4f8;
    border-right: 1px solid #d0d8e0;
}
.stTabs [aria-selected="true"] {
    background: #001c4b !important;
    color: #ffffff !important;
}

/* Boutons — angulaires, bleu ciel au hover */
.stButton > button {
    background: #001c4b;
    color: #ffffff;
    border: none;
    font-weight: 600;
    padding: 6px 15px;
    font-size: 0.80rem;
    transition: background 0.12s;
}
.stButton > button:hover { background: #019ee1; color: #ffffff; }

/* Bouton Sauvegarder — ciel (action principale, PAS orange) */
[data-testid="stFormSubmitButton"] > button {
    background: #019ee1 !important;
    color: #ffffff !important;
    border: none;
    font-weight: 700;
}
[data-testid="stFormSubmitButton"] > button:hover {
    background: #0180c0 !important;
}

/* Download buttons */
[data-testid="stDownloadButton"] > button {
    !important;
}

/* Inputs — angulaires */
.stSelectbox > div > div,
.stTextInput > div > div,
.stNumberInput > div > div,
.stMultiSelect > div {
    !important;
}

/* Labels */
.stSelectbox label, .stTextInput label, .stNumberInput label,
.stDateInput label, .stTextArea label, .stRadio label {
    color: #001c4b !important;
    font-weight: 600;
    font-size: 0.78rem;
}

h1, h2, h3, h4 { color: #001c4b !important; }
hr { border-color: #001c4b10; }
code { background: #001c4b08; color: #001c4b; }

/* Hint pipeline */
.pipeline-hint {
    background: #019ee108;
    border-left: 2px solid #001c4b;
    padding: 6px 11px;
    font-size: 0.77rem;
    color: #001c4b;
    margin-bottom: 8px;
}

/* Sidebar KPI blocks — sans arrondi */
.sidebar-kpi {
    background: #ffffff14;
    padding: 8px;
    margin-bottom: 5px;
}
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
# HELPERS SIMPLES
# ---------------------------------------------------------------------------

def fmt_m(v):
    """Formatage M EUR / Md EUR avec 1 decimale — standard Amundi."""
    try:
        fv = float(v)
    except (TypeError, ValueError):
        return "—"
    if fv >= 1_000_000_000:
        return "{:.1f} Md\u20ac".format(fv / 1_000_000_000)
    if fv >= 1_000_000:
        return "{:.1f} M\u20ac".format(fv / 1_000_000)
    if fv >= 1_000:
        return "{:.0f} k\u20ac".format(fv / 1_000)
    return "{:.0f} \u20ac".format(fv)


def kpi_card(label, value, sub=""):
    html = (
        '<div class="kpi-card">'
        '<div class="kpi-label">{}</div>'
        '<div class="kpi-value">{}</div>'
    ).format(label, value)
    if sub:
        html += '<div class="kpi-sub">{}</div>'.format(sub)
    html += "</div>"
    return html


def statut_badge(statut):
    colors = {
        "Funded":        (B_MID,    BLANC),
        "Soft Commit":   ("#2c7fb8", BLANC),
        "Due Diligence": ("#004f8c", BLANC),
        "Initial Pitch": (B_PAL,    BLANC),
        "Prospect":      ("#001c4b18", MARINE),
        "Lost":          (GRIS,     "#555"),
        "Paused":        ("#d0d0d0", MARINE),
        "Redeemed":      ("#c8d8e8", MARINE),
    }
    bg, fg = colors.get(statut, (GRIS, "#555"))
    return (
        '<span style="padding:2px 9px;'
        'font-size:0.71rem;font-weight:600;'
        'background:{};color:{};">{}</span>'
    ).format(bg, fg, statut)


# ---------------------------------------------------------------------------
# MODALES DRILL-DOWN — @st.dialog
# ---------------------------------------------------------------------------

@st.dialog("Detail — Deals Finances (Funded)")
def modal_funded(fonds_filter=None):
    df = db.get_funded_deals_detail(fonds_filter)
    if df.empty:
        st.info("Aucun deal Funded dans le perimetre.")
        return
    df2 = df.copy()
    if "AUM_Finance" in df2.columns:
        df2["AUM_Finance"] = df2["AUM_Finance"].apply(fmt_m)
    st.dataframe(df2, use_container_width=True, hide_index=True,
                 height=min(420, 46 + len(df2) * 36))


@st.dialog("Detail — Pipeline Actif")
def modal_pipeline_actif(fonds_filter=None):
    df = db.get_active_deals_detail(fonds_filter)
    if df.empty:
        st.info("Aucun deal actif dans le perimetre.")
        return
    df2 = df.copy()
    if "AUM_Revise" in df2.columns:
        df2["AUM_Revise"] = df2["AUM_Revise"].apply(fmt_m)
    if "Prochaine_Action" in df2.columns:
        df2["Prochaine_Action"] = df2["Prochaine_Action"].apply(
            lambda d: d.isoformat() if isinstance(d, date) else str(d or "—")
        )
    st.dataframe(df2, use_container_width=True, hide_index=True,
                 height=min(420, 46 + len(df2) * 36))


@st.dialog("Detail — Deals Perdus et En Pause")
def modal_lost(fonds_filter=None):
    df = db.get_lost_deals_detail(fonds_filter)
    if df.empty:
        st.info("Aucun deal Lost/Paused dans le perimetre.")
        return
    df2 = df.copy()
    if "AUM_Cible" in df2.columns:
        df2["AUM_Cible"] = df2["AUM_Cible"].apply(fmt_m)
    st.dataframe(df2, use_container_width=True, hide_index=True,
                 height=min(420, 46 + len(df2) * 36))


@st.dialog("Detail — Activites et Notes")
def modal_activities(client_nom=None):
    """Affiche toutes les activites + notes pour un client ou les 30 dernieres."""
    df = db.get_activities()
    if not df.empty and client_nom:
        df = df[df["nom_client"] == client_nom]
    if df.empty:
        st.info("Aucune activite enregistree.")
        return
    # Affichage plat, lisible, avec notes completes
    for _, row in df.iterrows():
        st.markdown(
            "<div style='border-left:2px solid #019ee1;padding:6px 12px;"
            "margin:4px 0;background:#f4f8fc;'>"
            "<div style='font-size:0.78rem;font-weight:700;color:#001c4b;'>"
            "{client} &nbsp;<span style='color:#019ee1;font-weight:400;'>"
            "{type}</span>&nbsp;"
            "<span style='color:#888;font-size:0.70rem;'>{date}</span></div>"
            "<div style='font-size:0.80rem;color:#444;margin-top:3px;'>"
            "{notes}</div></div>".format(
                client=row.get("nom_client",""),
                type=row.get("type_interaction",""),
                date=str(row.get("date","")),
                notes=str(row.get("notes","")) or "<em style='color:#aaa;'>Aucune note</em>",
            ),
            unsafe_allow_html=True
        )


@st.dialog("Detail — Actions en Retard")
def modal_overdue_detail():
    df = db.get_overdue_actions()
    if df.empty:
        st.info("Aucune action en retard.")
        return
    today = date.today()
    for _, row in df.iterrows():
        nad = row.get("next_action_date")
        days = (today - nad).days if isinstance(nad, date) else 0
        nad_str = nad.isoformat() if isinstance(nad, date) else "—"
        st.markdown(
            "<div style='border-left:2px solid #f07d00;padding:6px 12px;"
            "margin:4px 0;'>"
            "<div style='font-size:0.80rem;font-weight:700;color:#001c4b;'>"
            "{client}</div>"
            "<div style='font-size:0.75rem;color:#555;'>"
            "{fonds} &middot; {statut} &middot; "
            "<b>{owner}</b></div>"
            "<div style='font-size:0.72rem;margin-top:2px;'>"
            "Prevue le <b>{nad}</b> &nbsp;"
            "<span style='background:#f07d00;color:#fff;padding:1px 6px;"
            "font-size:0.66rem;font-weight:700;'>RETARD +{days}j</span>"
            "</div></div>".format(
                client=row.get("nom_client",""),
                fonds=row.get("fonds",""),
                statut=row.get("statut",""),
                owner=row.get("sales_owner",""),
                nad=nad_str, days=days,
            ),
            unsafe_allow_html=True
        )


# ---------------------------------------------------------------------------
# SIDEBAR
# ---------------------------------------------------------------------------

with st.sidebar:
    st.markdown(
        '<div style="text-align:center;padding:6px 0 14px 0;">'
        '<div style="font-size:0.91rem;font-weight:800;color:#ffffff;">'
        'CRM Asset Management</div>'
        '<div style="font-size:0.65rem;color:#e8e8e8;margin-top:3px;">'
        + date.today().strftime("%d %B %Y") +
        '</div></div>',
        unsafe_allow_html=True
    )
    st.divider()

    # Apercu global
    kpis_global = db.get_kpis()
    st.markdown(
        '<div style="color:#4a8fbd;font-size:0.67rem;font-weight:700;'
        'text-transform:uppercase;letter-spacing:1px;margin-bottom:5px;">'
        'Apercu Global</div>',
        unsafe_allow_html=True
    )
    for lbl, val in [
        ("AUM Finance Total", fmt_m(kpis_global["total_funded"])),
        ("Pipeline Actif",    fmt_m(kpis_global["pipeline_actif"])),
    ]:
        st.markdown(
            '<div class="sidebar-kpi">'
            '<div style="font-size:0.62rem;color:#c8dde8;">{}</div>'
            '<div style="font-size:1.06rem;font-weight:800;color:#ffffff;">{}</div>'
            '</div>'.format(lbl, val),
            unsafe_allow_html=True
        )

    st.divider()

    # Bouton backup Excel — securite si PDF/web plante
    st.markdown(
        '<div style="color:#4a8fbd;font-size:0.67rem;font-weight:700;'
        'text-transform:uppercase;letter-spacing:1px;margin-bottom:5px;">'
        'Securite Backup</div>',
        unsafe_allow_html=True
    )
    try:
        excel_bytes = db.get_excel_backup()
        st.download_button(
            label="Exporter Backup Excel",
            data=excel_bytes,
            file_name="backup_crm_{}.xlsx".format(date.today().isoformat()),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            help="Export .xlsx : Pipeline, Audit, Clients — disponible si le PDF plante.",
        )
    except Exception as e_xl:
        st.error("Backup Excel : {}".format(e_xl))

    st.divider()

    # Perimetre export PDF
    st.markdown(
        '<div style="color:#4a8fbd;font-size:0.67rem;font-weight:700;'
        'text-transform:uppercase;letter-spacing:1px;margin-bottom:5px;">'
        'Export PDF</div>',
        unsafe_allow_html=True
    )

    fonds_perimetre = st.multiselect(
        "Perimetre de l'export",
        options=FONDS,
        default=FONDS,
        key="fonds_perimetre_select",
        help="Fonds inclus dans le rapport PDF. Filtre tous les KPIs et graphiques."
    )

    _filtre_effectif = (
        fonds_perimetre
        if (fonds_perimetre and len(fonds_perimetre) < len(FONDS))
        else None
    )

    mode_comex = st.toggle(
        "Mode Comex - Anonymisation", value=False,
        help="Remplace les noms clients par Type - Region dans le PDF."
    )

    if fonds_perimetre and len(fonds_perimetre) < len(FONDS):
        badges = " ".join(
            '<span class="perimetre-badge">{}</span>'.format(f)
            for f in fonds_perimetre
        )
        st.markdown("<div>{}</div>".format(badges), unsafe_allow_html=True)
    elif not fonds_perimetre:
        st.warning("Selectionnez au moins un fonds.")

    perf_data_pdf   = st.session_state.get("perf_data", None)
    nav_base100_pdf = st.session_state.get("nav_base100", None)
    if perf_data_pdf is not None:
        st.caption("NAV disponible : {} fonds.".format(len(perf_data_pdf)))

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
                if pf_pdf is not None and _filtre_effectif:
                    if "Fonds" in pf_pdf.columns:
                        pf_pdf = pf_pdf[pf_pdf["Fonds"].isin(_filtre_effectif)]
                if nb_pdf is not None and _filtre_effectif:
                    cols_k = [c for c in nb_pdf.columns if c in _filtre_effectif]
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
                fname = "report{}_{}.pdf".format(
                    "_comex" if mode_comex else "",
                    date.today().isoformat()
                )
                st.download_button(
                    "Telecharger le rapport",
                    data=pdf_bytes, file_name=fname,
                    mime="application/pdf", use_container_width=True,
                )
                st.success("Rapport genere.")
            except Exception as e:
                st.error("Erreur generation PDF : {}".format(e))

    st.divider()
    st.caption("Version 5.1 — Amundi Edition Stable")


# ---------------------------------------------------------------------------
# EN-TETE
# ---------------------------------------------------------------------------

st.markdown(
    '<div class="crm-header">'
    '<h1>CRM &amp; Reporting — Asset Management</h1>'
    '<p>Pipeline commercial &middot; Suivi des mandats &middot;'
    ' Reporting executif &middot; Performance NAV</p>'
    '</div>',
    unsafe_allow_html=True
)


# ---------------------------------------------------------------------------
# ONGLETS
# ---------------------------------------------------------------------------

tab_ingest, tab_pipeline, tab_dash, tab_sales, tab_perf = st.tabs([
    "Data Ingestion",
    "Pipeline Management",
    "Executive Dashboard",
    "Sales Tracking",
    "Performance & NAV",
])


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
                sub_c = st.form_submit_button("Enregistrer le client",
                                              use_container_width=True)
            if sub_c:
                if not nom_client.strip():
                    st.error("Nom du client obligatoire.")
                else:
                    try:
                        db.add_client(nom_client.strip(), type_client, region)
                        st.success("Client {} ajoute.".format(nom_client))
                    except Exception as e:
                        st.warning("Ce client existe deja."
                                   if "UNIQUE" in str(e) else "Erreur : {}".format(e))

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
                        owner_opts  = sales_owners + ["Nouveau..."]
                        owner_sel   = st.selectbox("Commercial", owner_opts)
                        owner_input = (
                            st.text_input("Nom du nouveau commercial")
                            if owner_sel == "Nouveau..."
                            else owner_sel
                        )
                    with cb:
                        target_aum  = st.number_input("AUM Cible (EUR)", min_value=0, step=1_000_000)
                        revised_aum = st.number_input("AUM Revise (EUR)", min_value=0, step=1_000_000)
                        funded_aum  = st.number_input("AUM Finance (EUR)", min_value=0, step=1_000_000)

                    raison_perte, concurrent = "", ""
                    if statut_sel in ("Lost", "Paused"):
                        cc, cd = st.columns(2)
                        with cc:
                            raison_perte = st.selectbox("Raison", RAISONS_PERTE)
                        with cd:
                            concurrent = st.text_input("Concurrent")

                    next_action = st.date_input("Prochaine Action",
                                                value=date.today() + timedelta(days=14))
                    sub_d = st.form_submit_button("Enregistrer le deal",
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
                        st.success("Deal {}/{} enregistre.".format(fonds_sel, client_sel))

        with st.expander("Enregistrer une Activite"):
            clients_dict2 = db.get_client_options()
            if clients_dict2:
                with st.form("form_act", clear_on_submit=True):
                    ce, cf = st.columns(2)
                    with ce:
                        act_client = st.selectbox("Client", list(clients_dict2.keys()))
                        act_type   = st.selectbox("Type d'interaction", TYPES_INTERACTION)
                    with cf:
                        act_date  = st.date_input("Date", value=date.today())
                        act_notes = st.text_area("Notes", height=68)
                    sub_a = st.form_submit_button("Enregistrer", use_container_width=True)
                if sub_a:
                    db.add_activity(clients_dict2[act_client],
                                    act_date.isoformat(), act_notes, act_type)
                    st.success("Activite enregistree pour {}.".format(act_client))

    with col_import:
        st.markdown("#### Import CSV / Excel - Upsert")
        import_type = st.radio("Table cible", ["Clients", "Pipeline"], horizontal=True)
        if import_type == "Clients":
            st.info("Colonnes : nom_client, type_client, region")
        else:
            st.info("Colonnes : nom_client, fonds, statut, target_aum_initial,"
                    " revised_aum, funded_aum, raison_perte, concurrent_choisi,"
                    " next_action_date, sales_owner")

        uploaded_file = st.file_uploader("Fichier CSV ou Excel",
                                          type=["csv", "xlsx", "xls"])
        if uploaded_file:
            try:
                df_imp = (pd.read_csv(uploaded_file)
                          if uploaded_file.name.endswith(".csv")
                          else pd.read_excel(uploaded_file))
                st.dataframe(df_imp.head(5), use_container_width=True, height=145)
                st.caption("{} lignes detectees".format(len(df_imp)))
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
            for _, arow in df_act.head(10).iterrows():
                notes_preview = str(arow.get("notes") or "")[:60]
                if len(str(arow.get("notes") or "")) > 60:
                    notes_preview += "..."
                col_a, col_b = st.columns([4, 1])
                with col_a:
                    st.markdown(
                        "<div style='border-left:2px solid #019ee1;padding:4px 10px;"
                        "margin:2px 0;'>"
                        "<span style='font-size:0.78rem;font-weight:700;color:#001c4b;'>"
                        "{client}</span>"
                        " <span style='font-size:0.70rem;color:#019ee1;'>{type}</span>"
                        " <span style='font-size:0.68rem;color:#888;'>{date}</span>"
                        "<div style='font-size:0.74rem;color:#555;margin-top:1px;'>"
                        "{notes}</div></div>".format(
                            client=arow.get("nom_client",""),
                            type=arow.get("type_interaction",""),
                            date=str(arow.get("date","")),
                            notes=notes_preview or "<em>—</em>",
                        ),
                        unsafe_allow_html=True
                    )
                with col_b:
                    if st.button("Notes", key="act_{}".format(arow.get("id","")),
                                 use_container_width=True):
                        modal_activities(arow.get("nom_client"))


# ============================================================================
# ONGLET 2 — PIPELINE MANAGEMENT
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
        '<div class="pipeline-hint">Selectionnez une ligne pour modifier'
        ' — <b>{} deal(s)</b> affiches</div>'.format(len(df_view)),
        unsafe_allow_html=True
    )

    df_display = df_view.copy()
    df_display["next_action_date"] = df_display["next_action_date"].apply(
        lambda d: d.isoformat() if isinstance(d, date) else ""
    )
    cols_show = ["id", "nom_client", "type_client", "region", "fonds", "statut",
                 "target_aum_initial", "revised_aum", "funded_aum",
                 "raison_perte", "next_action_date", "sales_owner"]

    event = st.dataframe(
        df_display[cols_show],
        use_container_width=True, height=380, hide_index=True,
        on_select="rerun", selection_mode="single-row",
        column_config={
            "id":                 st.column_config.NumberColumn("ID", width="small"),
            "nom_client":         st.column_config.TextColumn("Client"),
            "type_client":        st.column_config.TextColumn("Type", width="small"),
            "region":             st.column_config.TextColumn("Region", width="small"),
            "fonds":              st.column_config.TextColumn("Fonds"),
            "statut":             st.column_config.TextColumn("Statut"),
            "target_aum_initial": st.column_config.NumberColumn("AUM Cible", format="%.0f"),
            "revised_aum":        st.column_config.NumberColumn("AUM Revise", format="%.0f"),
            "funded_aum":         st.column_config.NumberColumn("AUM Finance", format="%.0f"),
            "raison_perte":       st.column_config.TextColumn("Raison"),
            "next_action_date":   st.column_config.TextColumn("Next Action"),
            "sales_owner":        st.column_config.TextColumn("Commercial"),
        },
        key="pipeline_ro"
    )

    selected_rows = event.selection.rows if event.selection else []
    pipeline_id   = None

    if len(selected_rows) > 0:
        sel_idx = selected_rows[0]
        # Use df_display (not df_view) — it matches what is shown in the table.
        # df_view may have a different length if filters changed after rendering.
        if sel_idx < len(df_display):
            pipeline_id = int(df_display.iloc[sel_idx]["id"])

    if pipeline_id is not None:
        row_data = db.get_pipeline_row_by_id(pipeline_id)

        if row_data:
            client_name    = str(row_data.get("nom_client", ""))
            current_statut = str(row_data.get("statut", "Prospect"))

            st.markdown(
                '<div class="detail-panel">'
                '<div style="font-size:0.88rem;font-weight:700;color:{marine};">'
                'Modification — <span style="color:{ciel};">{name}</span>'
                '&nbsp; {badge}'
                '&nbsp;<span style="font-size:0.68rem;color:#888;font-weight:400;">'
                'ID #{pid}</span></div></div>'.format(
                    marine=MARINE, ciel=B_MID,
                    name=client_name,
                    badge=statut_badge(current_statut),
                    pid=pipeline_id
                ),
                unsafe_allow_html=True
            )

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
                                   if new_statut in ("Lost", "Paused")
                                   else "Raison Perte/Pause")
                        new_raison = st.selectbox(lbl_r, ropts, index=ri)
                    with r3c2:
                        new_conc = st.text_input(
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

                    sub = st.form_submit_button("Sauvegarder les modifications",
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
                        "concurrent_choisi":  new_conc,
                        "next_action_date":   new_nad,
                        "sales_owner":        new_sales,
                    })
                    if ok:
                        st.success("Deal mis a jour — Statut : {} / AUM Finance : {}".format(
                            new_statut, fmt_m(new_funded)))
                        st.rerun()
                    else:
                        st.error(msg)

            # Audit trail sobre
            df_audit = db.get_audit_log(pipeline_id)
            if not df_audit.empty:
                with st.expander(
                    "Historique des modifications — Deal #{}".format(pipeline_id),
                    expanded=False
                ):
                    st.dataframe(df_audit, use_container_width=True, hide_index=True,
                                 height=min(220, 46 + len(df_audit) * 36))
            else:
                st.caption("Aucune modification enregistree pour ce deal.")
    else:
        st.markdown(
            '<div style="background:#001c4b04;border:1px dashed #001c4b20;'
            'padding:20px;text-align:center;margin-top:10px;">'
            '<div style="color:{};font-weight:600;font-size:0.85rem;">'
            'Selectionnez un deal dans le tableau</div>'
            '<div style="color:#888;font-size:0.75rem;margin-top:2px;">'
            'Le formulaire d\'edition et l\'historique s\'afficheront ici</div>'
            '</div>'.format(MARINE),
            unsafe_allow_html=True
        )

    # Graphique Plotly — AUM par deal actif — EXPLICITE, zero helper dict
    st.divider()
    st.markdown("#### Comparaison AUM par Deal Actif")
    df_viz = df_pipe[
        (df_pipe["target_aum_initial"] > 0) &
        (df_pipe["statut"].isin(["Funded", "Soft Commit", "Due Diligence", "Initial Pitch"]))
    ].copy().head(10)

    if not df_viz.empty:
        clients_viz = df_viz["nom_client"].str[:18].tolist()
        target_vals = df_viz["target_aum_initial"].tolist()
        revised_vals= df_viz["revised_aum"].tolist()
        funded_vals = df_viz["funded_aum"].tolist()

        fig_viz = go.Figure()
        fig_viz.add_trace(go.Bar(
            name="AUM Cible",
            x=clients_viz,
            y=target_vals,
            marker_color=GRIS,
            marker_line_color=BLANC,
            marker_line_width=0.5,
            hovertemplate="<b>%{x}</b><br>AUM Cible : %{y:,.0f} EUR<extra></extra>",
        ))
        fig_viz.add_trace(go.Bar(
            name="AUM Revise",
            x=clients_viz,
            y=revised_vals,
            marker_color=B_MID,
            marker_line_color=BLANC,
            marker_line_width=0.5,
            hovertemplate="<b>%{x}</b><br>AUM Revise : %{y:,.0f} EUR<extra></extra>",
        ))
        fig_viz.add_trace(go.Bar(
            name="AUM Finance",
            x=clients_viz,
            y=funded_vals,
            marker_color=MARINE,
            marker_line_color=BLANC,
            marker_line_width=0.5,
            hovertemplate="<b>%{x}</b><br>AUM Finance : %{y:,.0f} EUR<extra></extra>",
        ))
        fig_viz.update_layout(
            title_text="AUM Cible / Revise / Finance — Deals Actifs",
            title_font_size=13,
            title_font_color=MARINE,
            height=340,
            paper_bgcolor=BLANC,
            plot_bgcolor=BLANC,
            font_color=MARINE,
            font_family="Segoe UI, Arial",
            barmode="group",
            bargap=0.22,
            bargroupgap=0.06,
            legend_bgcolor=BLANC,
            legend_bordercolor=GRIS,
            legend_borderwidth=1,
            legend_font_size=10,
            xaxis_tickangle=-20,
            xaxis_tickfont_size=10,
            xaxis_tickfont_color=MARINE,
            xaxis_showgrid=False,
            yaxis_showgrid=True,
            yaxis_gridcolor=GRIS,
            yaxis_tickfont_size=9,
            yaxis_tickfont_color=MARINE,
            margin_l=10, margin_r=10, margin_t=40, margin_b=10,
        )
        st.plotly_chart(fig_viz, use_container_width=True,
                        config={"displayModeBar": False})

    # Tableau Lost / Paused
    df_lp = df_pipe[df_pipe["statut"].isin(["Lost", "Paused"])].copy()
    if not df_lp.empty:
        st.divider()
        st.markdown("#### Deals Perdus / En Pause")
        df_lp2 = df_lp[["nom_client", "fonds", "statut", "target_aum_initial",
                          "raison_perte", "concurrent_choisi", "sales_owner"]].copy()
        df_lp2["target_aum_initial"] = df_lp2["target_aum_initial"].apply(fmt_m)
        st.dataframe(df_lp2, use_container_width=True, hide_index=True,
                     column_config={
                         "nom_client":         st.column_config.TextColumn("Client"),
                         "fonds":              st.column_config.TextColumn("Fonds"),
                         "statut":             st.column_config.TextColumn("Statut"),
                         "target_aum_initial": st.column_config.TextColumn("AUM Cible"),
                         "raison_perte":       st.column_config.TextColumn("Raison"),
                         "concurrent_choisi":  st.column_config.TextColumn("Concurrent"),
                         "sales_owner":        st.column_config.TextColumn("Commercial"),
                     })


# ============================================================================
# ONGLET 3 — EXECUTIVE DASHBOARD
# ============================================================================
with tab_dash:
    st.markdown('<div class="section-title">Executive Dashboard</div>',
                unsafe_allow_html=True)

    kpis = db.get_kpis()
    nb_lost_paused = kpis["nb_lost"] + kpis.get("nb_paused", 0)

    # KPI Cards — interactivite via boutons Streamlit integres dans le bloc visuel
    # Le bouton est rendu sous la card mais visuellement "fusionne" via CSS
    kc1, kc2, kc3, kc4 = st.columns(4, gap="small")

    with kc1:
        st.markdown(
            '<div class="kpi-card" style="border-bottom:2px solid #019ee1;">'
            '<div class="kpi-label">AUM Finance Total</div>'
            '<div class="kpi-value">{}</div>'
            '<div class="kpi-sub" style="color:#019ee1;cursor:pointer;">'
            '{} deal(s) Funded &rsaquo;</div>'
            '</div>'.format(fmt_m(kpis["total_funded"]), kpis["nb_funded"]),
            unsafe_allow_html=True
        )
        if st.button("Detail AUM Finance", key="btn_funded",
                     use_container_width=True, type="secondary"):
            modal_funded(_filtre_effectif)

    with kc2:
        st.markdown(
            '<div class="kpi-card" style="border-bottom:2px solid #019ee1;">'
            '<div class="kpi-label">Pipeline Actif</div>'
            '<div class="kpi-value">{}</div>'
            '<div class="kpi-sub" style="color:#019ee1;cursor:pointer;">'
            '{} deals en cours &rsaquo;</div>'
            '</div>'.format(fmt_m(kpis["pipeline_actif"]), kpis["nb_deals_actifs"]),
            unsafe_allow_html=True
        )
        if st.button("Detail Pipeline", key="btn_pipe",
                     use_container_width=True, type="secondary"):
            modal_pipeline_actif(_filtre_effectif)

    with kc3:
        st.markdown(
            '<div class="kpi-card">'
            '<div class="kpi-label">Taux Conversion</div>'
            '<div class="kpi-value">{:.1f}%</div>'
            '<div class="kpi-sub">{} funded / {} lost</div>'
            '</div>'.format(
                kpis["taux_conversion"], kpis["nb_funded"], kpis["nb_lost"]),
            unsafe_allow_html=True
        )

    with kc4:
        st.markdown(
            '<div class="kpi-card" style="border-bottom:2px solid #019ee1;">'
            '<div class="kpi-label">Lost / Paused</div>'
            '<div class="kpi-value">{}</div>'
            '<div class="kpi-sub" style="color:#019ee1;cursor:pointer;">'
            'Analyser &rsaquo;</div>'
            '</div>'.format(nb_lost_paused),
            unsafe_allow_html=True
        )
        if nb_lost_paused > 0:
            if st.button("Detail Lost / Paused", key="btn_lost",
                         use_container_width=True, type="secondary"):
                modal_lost(_filtre_effectif)

    st.markdown("<br>", unsafe_allow_html=True)

    # Pastilles de statut
    statut_order = [s for s in STATUTS if kpis["statut_repartition"].get(s, 0) > 0]
    if statut_order:
        bcols = st.columns(len(statut_order), gap="small")
        for col, s in zip(bcols, statut_order):
            c_hex = STATUT_COLORS.get(s, GRIS)
            with col:
                st.markdown(
                    '<div style="background:{c}16;border:1px solid {c}44;'
                    'padding:7px;text-align:center;">'
                    '<div style="font-size:0.61rem;color:{marine};font-weight:700;'
                    'text-transform:uppercase;">{s}</div>'
                    '<div style="font-size:1.3rem;font-weight:800;color:{c};">{n}</div>'
                    '</div>'.format(
                        c=c_hex, marine=MARINE,
                        s=s, n=kpis["statut_repartition"][s]
                    ),
                    unsafe_allow_html=True
                )

    st.divider()

    # Alertes retard
    df_overdue = db.get_overdue_actions()
    if not df_overdue.empty:
        col_ov, col_ov_btn = st.columns([6, 1])
        with col_ov:
            st.markdown(
                "<div style='font-size:0.84rem;font-weight:700;color:#001c4b;"
                "border-left:3px solid #f07d00;padding:4px 10px;'>"
                "{} action(s) en retard</div>".format(len(df_overdue)),
                unsafe_allow_html=True
            )
        with col_ov_btn:
            if st.button("Voir detail", key="btn_overdue", use_container_width=True):
                modal_overdue_detail()
        for _, row in df_overdue.iterrows():
            nad = row.get("next_action_date")
            if isinstance(nad, date):
                days_late = (date.today() - nad).days
                nad_str   = nad.isoformat()
            else:
                days_late = 0
                nad_str   = str(nad or "—")
            owner = str(row.get("sales_owner", "")) or ""
            st.markdown(
                '<div class="alert-overdue">'
                '<b>{client}</b> — {fonds}'
                ' <span style="color:{ciel};font-weight:600;">({statut})</span>'
                ' — Prevue le <b>{nad}</b>'
                ' &nbsp;<span class="badge-retard">RETARD +{days}j</span>'
                '{owner_part}'
                '</div>'.format(
                    client=row["nom_client"], fonds=row["fonds"],
                    ciel=CIEL, statut=row["statut"], nad=nad_str,
                    days=days_late,
                    owner_part=" — <b>{}</b>".format(owner) if owner else ""
                ),
                unsafe_allow_html=True
            )

    # Graphiques Plotly — EXPLICITES, zero helper dict
    gcol1, gcol2, gcol3 = st.columns([1, 1, 1.2], gap="medium")

    # --- Donut Type Client ---
    with gcol1:
        st.markdown("#### Par Type Client")
        abt = kpis.get("aum_by_type", {})
        if abt:
            labels_t = list(abt.keys())
            values_t = list(abt.values())
            fig_type = go.Figure(go.Pie(
                labels=labels_t,
                values=values_t,
                hole=0.52,
                marker_colors=PALETTE[:len(labels_t)],
                marker_line_color=BLANC,
                marker_line_width=2,
                hovertemplate="<b>%{label}</b><br>%{percent}<extra></extra>",
                textinfo="percent",
                textfont_size=10,
                textfont_color=BLANC,
                insidetextorientation="auto",
            ))
            fig_type.add_annotation(
                text=fmt_m(sum(values_t)),
                x=0.5, y=0.55,
                font_size=11, font_color=MARINE, font_family="Segoe UI",
                showarrow=False
            )
            fig_type.add_annotation(
                text="Finance",
                x=0.5, y=0.42,
                font_size=8, font_color=GTXT,
                showarrow=False
            )
            fig_type.update_layout(
                title_text="AUM par Type",
                title_font_size=12,
                title_font_color=MARINE,
                height=300,
                paper_bgcolor=BLANC,
                plot_bgcolor=BLANC,
                font_color=MARINE,
                font_family="Segoe UI, Arial",
                showlegend=True,
                legend_orientation="v",
                legend_x=1.02,
                legend_y=0.5,
                legend_font_size=9,
                legend_font_color=MARINE,
                margin_l=0, margin_r=80, margin_t=36, margin_b=10,
            )
            st.plotly_chart(fig_type, use_container_width=True,
                            config={"displayModeBar": False})

    # --- Donut Region ---
    with gcol2:
        st.markdown("#### Par Region")
        aum_reg_dash = db.get_aum_by_region()
        if aum_reg_dash:
            labels_r = list(aum_reg_dash.keys())
            values_r = list(aum_reg_dash.values())
            fig_reg = go.Figure(go.Pie(
                labels=labels_r,
                values=values_r,
                hole=0.52,
                marker_colors=PALETTE[:len(labels_r)],
                marker_line_color=BLANC,
                marker_line_width=2,
                hovertemplate="<b>%{label}</b><br>%{percent}<extra></extra>",
                textinfo="percent",
                textfont_size=10,
                textfont_color=BLANC,
                insidetextorientation="auto",
            ))
            fig_reg.add_annotation(
                text=fmt_m(sum(values_r)),
                x=0.5, y=0.55,
                font_size=11, font_color=MARINE, font_family="Segoe UI",
                showarrow=False
            )
            fig_reg.add_annotation(
                text="Finance",
                x=0.5, y=0.42,
                font_size=8, font_color=GTXT,
                showarrow=False
            )
            fig_reg.update_layout(
                title_text="AUM par Region",
                title_font_size=12,
                title_font_color=MARINE,
                height=300,
                paper_bgcolor=BLANC,
                plot_bgcolor=BLANC,
                font_color=MARINE,
                font_family="Segoe UI, Arial",
                showlegend=True,
                legend_orientation="v",
                legend_x=1.02,
                legend_y=0.5,
                legend_font_size=9,
                legend_font_color=MARINE,
                margin_l=0, margin_r=80, margin_t=36, margin_b=10,
            )
            st.plotly_chart(fig_reg, use_container_width=True,
                            config={"displayModeBar": False})

    # --- Bar AUM par Fonds ---
    with gcol3:
        st.markdown("#### AUM par Fonds")
        abf = kpis.get("aum_by_fonds", {})
        if abf:
            fonds_sorted = sorted(abf.items(), key=lambda x: x[1], reverse=True)
            f_lbls = [f[0] for f in fonds_sorted]
            f_vals = [f[1] for f in fonds_sorted]
            f_text = [fmt_m(v) for v in f_vals]
            fig_fonds = go.Figure(go.Bar(
                x=f_vals,
                y=f_lbls,
                orientation="h",
                marker_color=PALETTE[:len(f_lbls)],
                marker_line_color=BLANC,
                marker_line_width=0.5,
                hovertemplate="<b>%{y}</b><br>%{x:,.0f} EUR<extra></extra>",
                text=f_text,
                textposition="outside",
                textfont_size=9,
                textfont_color=MARINE,
            ))
            fig_fonds.update_layout(
                title_text="AUM Finance par Fonds",
                title_font_size=12,
                title_font_color=MARINE,
                height=300,
                paper_bgcolor=BLANC,
                plot_bgcolor=BLANC,
                font_color=MARINE,
                font_family="Segoe UI, Arial",
                xaxis_showgrid=True,
                xaxis_gridcolor=GRIS,
                xaxis_tickfont_size=8,
                xaxis_tickfont_color=MARINE,
                yaxis_autorange="reversed",
                yaxis_tickfont_size=9,
                yaxis_tickfont_color=MARINE,
                margin_l=10, margin_r=60, margin_t=36, margin_b=10,
            )
            st.plotly_chart(fig_fonds, use_container_width=True,
                            config={"displayModeBar": False})

    st.divider()

    # Top Deals
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
            bar_c = MARINE if i == 0 else B_MID if i < 3 else B_PAL
            st.markdown(
                '<div style="display:flex;align-items:center;gap:10px;'
                'margin:4px 0;padding:7px 12px;background:transparent;'
                'border-bottom:1px solid #e8e8e8;">'
                '<div style="font-size:0.72rem;font-weight:700;'
                'color:{ciel};min-width:26px;">No.{rank}</div>'
                '<div style="flex:1;">'
                '<div style="font-size:0.80rem;font-weight:700;'
                'color:{marine};">{client}</div>'
                '<div style="font-size:0.67rem;color:#777;">'
                '{fonds} &middot; {type} &middot; {region}</div>'
                '<div style="background:{gris};height:3px;'
                'margin-top:4px;overflow:hidden;">'
                '<div style="background:{barc};width:{pct:.0f}%;'
                'height:100%;"></div></div></div>'
                '<div style="text-align:right;min-width:72px;">'
                '<div style="font-size:0.86rem;font-weight:800;'
                'color:{marine};">{aum}</div>'
                '<div style="font-size:0.61rem;color:#999;">finance</div>'
                '</div></div>'.format(
                    rank=i + 1, ciel=CIEL, marine=MARINE, gris=GRIS,
                    client=row["nom_client"], fonds=row["fonds"],
                    type=row["type_client"], region=row["region"],
                    barc=bar_c, pct=pct, aum=fmt_m(val)
                ),
                unsafe_allow_html=True
            )


# ============================================================================
# ONGLET 4 — SALES TRACKING
# ============================================================================
with tab_sales:
    st.markdown(
        '<div class="section-title">Sales Tracking — Suivi par Commercial</div>',
        unsafe_allow_html=True
    )

    df_sm = db.get_sales_metrics()
    df_na = db.get_next_actions_by_sales(days_ahead=30)

    if df_sm.empty:
        st.info("Aucune donnee de pipeline disponible.")
    else:
        n_own  = len(df_sm)
        n_cols = min(n_own, 3)
        s_cols = st.columns(n_cols, gap="medium")

        for i, (_, row) in enumerate(df_sm.iterrows()):
            retard_val = int(row.get("Retards", 0))
            if retard_val > 0:
                retard_html = '<span class="badge-retard">RETARD : {}</span>'.format(retard_val)
            else:
                retard_html = (
                    '<span style="color:{};font-size:0.74rem;'
                    'font-weight:600;">A jour</span>'.format(CIEL)
                )
            with s_cols[i % n_cols]:
                st.markdown(
                    '<div class="sales-card">'
                    '<div class="sales-card-name">{name}</div>'
                    '<div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;">'
                    '<div><div class="sales-metric">Total Deals</div>'
                    '<div class="sales-metric-val">{nb}</div></div>'
                    '<div><div class="sales-metric">Funded</div>'
                    '<div class="sales-metric-val">{funded}</div></div>'
                    '<div><div class="sales-metric">AUM Finance</div>'
                    '<div class="sales-metric-acc">{aum}</div></div>'
                    '<div><div class="sales-metric">Pipeline Actif</div>'
                    '<div class="sales-metric-val">{pipe}</div></div>'
                    '<div><div class="sales-metric">Actifs / Perdus</div>'
                    '<div class="sales-metric-val">{actifs} / {perdus}</div></div>'
                    '<div><div class="sales-metric">Alertes</div>'
                    '<div style="margin-top:2px;">{retard}</div></div>'
                    '</div></div><br>'.format(
                        name=row["Commercial"],
                        nb=int(row["Nb_Deals"]),
                        funded=int(row["Funded"]),
                        aum=fmt_m(float(row["AUM_Finance"])),
                        pipe=fmt_m(float(row["Pipeline_Actif"])),
                        actifs=int(row["Actifs"]),
                        perdus=int(row["Perdus"]),
                        retard=retard_html,
                    ),
                    unsafe_allow_html=True
                )

        st.divider()
        st.markdown("#### AUM Finance vs Pipeline par Commercial")

        # Graphique Plotly — EXPLICITE
        if df_sm["AUM_Finance"].sum() > 0:
            owners_l   = df_sm["Commercial"].tolist()
            aum_v      = df_sm["AUM_Finance"].tolist()
            pipe_v     = df_sm["Pipeline_Actif"].tolist()

            fig_sales = go.Figure()
            fig_sales.add_trace(go.Bar(
                name="AUM Finance",
                x=owners_l,
                y=aum_v,
                marker_color=MARINE,
                marker_line_color=BLANC,
                marker_line_width=0.5,
                hovertemplate="<b>%{x}</b><br>AUM Finance : %{y:,.0f} EUR<extra></extra>",
            ))
            fig_sales.add_trace(go.Bar(
                name="Pipeline Actif",
                x=owners_l,
                y=pipe_v,
                marker_color=B_MID,
                marker_line_color=BLANC,
                marker_line_width=0.5,
                hovertemplate="<b>%{x}</b><br>Pipeline Actif : %{y:,.0f} EUR<extra></extra>",
            ))
            fig_sales.update_layout(
                height=320,
                paper_bgcolor=BLANC,
                plot_bgcolor=BLANC,
                font_color=MARINE,
                font_family="Segoe UI, Arial",
                barmode="group",
                bargap=0.25,
                bargroupgap=0.08,
                legend_bgcolor=BLANC,
                legend_bordercolor=GRIS,
                legend_borderwidth=1,
                legend_font_size=10,
                xaxis_tickfont_size=10,
                xaxis_tickfont_color=MARINE,
                xaxis_showgrid=False,
                yaxis_showgrid=True,
                yaxis_gridcolor=GRIS,
                yaxis_tickfont_size=9,
                yaxis_tickfont_color=MARINE,
                margin_l=10, margin_r=10, margin_t=20, margin_b=10,
            )
            st.plotly_chart(fig_sales, use_container_width=True,
                            config={"displayModeBar": False})

        st.divider()
        st.markdown("#### Tableau de Bord Commercial")
        df_sd = df_sm.copy()
        df_sd["AUM_Finance"]    = df_sd["AUM_Finance"].apply(fmt_m)
        df_sd["Pipeline_Actif"] = df_sd["Pipeline_Actif"].apply(fmt_m)
        st.dataframe(df_sd, use_container_width=True, hide_index=True,
                     height=min(320, 50 + len(df_sd) * 40))

        st.divider()
        st.markdown("#### Prochaines Actions — 30 jours")
        if df_na.empty:
            st.info("Aucune action planifiee dans les 30 prochains jours.")
        else:
            owners_na = ["Tous"] + sorted(df_na["sales_owner"].dropna().unique().tolist())
            filter_o  = st.selectbox("Filtrer par commercial", owners_na, key="sel_na")
            df_nav    = df_na if filter_o == "Tous" else df_na[df_na["sales_owner"] == filter_o]
            today_d   = date.today()

            for _, row in df_nav.iterrows():
                nad = row.get("next_action_date")
                if isinstance(nad, date):
                    days_diff = (nad - today_d).days
                    nad_str   = nad.isoformat()
                    if days_diff <= 3:
                        urgence = '<span class="badge-retard">URGENT J+{}</span>'.format(days_diff)
                    elif days_diff <= 7:
                        urgence = (
                            '<span style="background:{marine};color:{blanc};'
                            'padding:1px 6px;'
                            'font-size:0.66rem;font-weight:700;">J+{d}</span>'.format(
                                marine=MARINE, blanc=BLANC, d=days_diff
                            )
                        )
                    else:
                        urgence = (
                            '<span style="background:{gris};color:{marine};'
                            'padding:1px 6px;'
                            'font-size:0.66rem;font-weight:600;">J+{d}</span>'.format(
                                gris=GRIS, marine=MARINE, d=days_diff
                            )
                        )
                else:
                    nad_str = "—"
                    urgence = ""

                aum_str = fmt_m(float(row.get("revised_aum", 0) or 0))
                st.markdown(
                    '<div style="display:flex;align-items:center;gap:10px;'
                    'padding:6px 12px;margin:3px 0;background:transparent;'
                    'border-bottom:1px solid #e8e8e8;'
                    'font-size:0.78rem;">'
                    '<div style="flex:1;">'
                    '<span style="font-weight:700;color:{marine};">{client}</span>'
                    ' <span style="color:#777;">— {fonds}</span>'
                    ' <span style="color:{ciel};font-size:0.71rem;">({statut})</span>'
                    '</div>'
                    '<div style="color:#555;min-width:78px;font-size:0.74rem;">{nad}</div>'
                    '<div style="min-width:58px;">{urgence}</div>'
                    '<div style="color:{mid};font-weight:700;min-width:72px;'
                    'text-align:right;">{aum}</div>'
                    '<div style="color:#888;font-size:0.71rem;min-width:78px;">{owner}</div>'
                    '</div>'.format(
                        marine=MARINE, ciel=CIEL, mid=B_MID,
                        client=row.get("nom_client", ""),
                        fonds=row.get("fonds", ""),
                        statut=row.get("statut", ""),
                        nad=nad_str, urgence=urgence,
                        aum=aum_str, owner=row.get("sales_owner", ""),
                    ),
                    unsafe_allow_html=True
                )


# ============================================================================
# ONGLET 5 — PERFORMANCE & NAV
# ============================================================================
with tab_perf:
    st.markdown(
        '<div class="section-title">Performance & NAV — Module Analytique</div>',
        unsafe_allow_html=True
    )

    col_up, col_demo = st.columns([2, 1], gap="medium")
    with col_up:
        nav_file = st.file_uploader(
            "Charger un fichier NAV (Excel ou CSV)",
            type=["xlsx", "xls", "csv"],
            help="Colonnes : Date, Fonds, NAV"
        )
    with col_demo:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Generer fichier NAV demo", use_container_width=True):
            today_ts = pd.Timestamp(date.today())
            dates_d  = pd.date_range(end=today_ts, periods=252, freq="B")
            demo_rows = []
            for fonds_d, seed in zip(FONDS, range(len(FONDS))):
                rng   = np.random.default_rng(seed=seed + 42)
                ret   = rng.normal(0.0003, 0.008, size=len(dates_d))
                nav_d = 100.0 * np.cumprod(1 + ret)
                for d, n in zip(dates_d, nav_d):
                    demo_rows.append({"Date": d.strftime("%Y-%m-%d"),
                                      "Fonds": fonds_d, "NAV": round(float(n), 6)})
            df_demo = pd.DataFrame(demo_rows)
            buf_demo = io.BytesIO()
            df_demo.to_excel(buf_demo, index=False, engine="openpyxl")
            st.download_button(
                "Telecharger le fichier demo",
                data=buf_demo.getvalue(),
                file_name="nav_demo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    if nav_file is not None:
        try:
            df_nav_raw = (pd.read_csv(nav_file)
                          if nav_file.name.endswith(".csv")
                          else pd.read_excel(nav_file))

            df_nav_raw.columns = [str(c).strip() for c in df_nav_raw.columns]
            col_map = {}
            for c in df_nav_raw.columns:
                cl = c.lower()
                if "date" in cl:             col_map[c] = "Date"
                elif "fonds" in cl or "fund" in cl: col_map[c] = "Fonds"
                elif "nav" in cl:            col_map[c] = "NAV"
            df_nav_raw = df_nav_raw.rename(columns=col_map)

            missing = {"Date", "Fonds", "NAV"} - set(df_nav_raw.columns)
            if missing:
                st.error("Colonnes manquantes : {}. Colonnes trouvees : {}".format(
                    ", ".join(missing), list(df_nav_raw.columns)))
            else:
                df_nav_raw["Date"] = pd.to_datetime(df_nav_raw["Date"], errors="coerce")
                df_nav_raw["NAV"]  = pd.to_numeric(df_nav_raw["NAV"], errors="coerce")
                df_nav_raw = df_nav_raw.dropna(subset=["Date", "NAV"]).copy()

                pf1, pf2, pf3 = st.columns(3)
                fonds_nav_opts = sorted(df_nav_raw["Fonds"].unique().tolist())
                with pf1:
                    fonds_sel_nav = st.multiselect("Fonds", fonds_nav_opts,
                                                    default=fonds_nav_opts, key="nav_fs")
                with pf2:
                    d_debut = st.date_input("Date debut",
                                             value=df_nav_raw["Date"].min().date(),
                                             key="nav_dd")
                with pf3:
                    d_fin = st.date_input("Date fin",
                                          value=df_nav_raw["Date"].max().date(),
                                          key="nav_df")

                mask = (
                    df_nav_raw["Fonds"].isin(fonds_sel_nav) &
                    (df_nav_raw["Date"] >= pd.Timestamp(d_debut)) &
                    (df_nav_raw["Date"] <= pd.Timestamp(d_fin))
                )
                df_nav_f = df_nav_raw[mask].copy()

                if df_nav_f.empty:
                    st.warning("Aucune donnee pour ces filtres.")
                else:
                    pivot = df_nav_f.pivot_table(
                        index="Date", columns="Fonds", values="NAV", aggfunc="last"
                    ).sort_index()

                    base100 = pivot.copy()
                    for col_b in base100.columns:
                        first_v = base100[col_b].dropna()
                        if not first_v.empty and float(first_v.iloc[0]) != 0:
                            base100[col_b] = base100[col_b] / float(first_v.iloc[0]) * 100

                    st.session_state["nav_base100"] = base100
                    st.session_state["perf_data"]   = None

                    # Graphique Base 100 — Plotly EXPLICITE
                    st.markdown("#### Courbe NAV — Base 100")
                    # Palette NAV — 6 couleurs vraiment distinctives, lisibles
                    # Principe : contraste maximal entre chaque serie
                    NAV_COLORS = [
                        "#001c4b",   # Marine profond
                        "#019ee1",   # Ciel Amundi
                        "#e63946",   # Rouge vif
                        "#2a9d8f",   # Vert teal
                        "#f4a261",   # Ambre chaud
                        "#9b59b6",   # Violet
                        "#27ae60",   # Vert emeraude
                        "#e67e22",   # Orange brun
                    ]

                    fig_nav = go.Figure()
                    for i_n, fonds_n in enumerate(base100.columns):
                        series_n = base100[fonds_n].dropna()
                        if series_n.empty:
                            continue
                        fig_nav.add_trace(go.Scatter(
                            x=series_n.index,
                            y=series_n.values,
                            name=fonds_n,
                            mode="lines",
                            line_color=NAV_COLORS[i_n % len(NAV_COLORS)],
                            line_width=2.0,
                            line_dash="solid",
                            hovertemplate=(
                                "<b>" + fonds_n + "</b><br>"
                                "Date : %{x|%d/%m/%Y}<br>"
                                "Base 100 : %{y:.2f}<extra></extra>"
                            ),
                        ))

                    fig_nav.add_hline(
                        y=100,
                        line_color=GRIS,
                        line_width=0.8,
                        line_dash="dot",
                    )
                    fig_nav.update_layout(
                        title_text="Performance NAV — Base 100",
                        title_font_size=13,
                        title_font_color=MARINE,
                        height=380,
                        paper_bgcolor=BLANC,
                        plot_bgcolor=BLANC,
                        font_color=MARINE,
                        font_family="Segoe UI, Arial",
                        hovermode="x unified",
                        legend_bgcolor=BLANC,
                        legend_bordercolor=GRIS,
                        legend_borderwidth=1,
                        legend_font_size=10,
                        xaxis_showgrid=False,
                        xaxis_tickfont_size=9,
                        xaxis_tickfont_color=MARINE,
                        yaxis_showgrid=True,
                        yaxis_gridcolor=GRIS,
                        yaxis_tickfont_size=9,
                        yaxis_tickfont_color=MARINE,
                        yaxis_zeroline=False,
                        margin_l=10, margin_r=10, margin_t=40, margin_b=10,
                    )
                    st.plotly_chart(fig_nav, use_container_width=True,
                                    config={"displayModeBar": True,
                                            "modeBarButtonsToRemove": ["lasso2d", "select2d"]})

                    # Calcul performances
                    today_ts  = pd.Timestamp(date.today())
                    one_m_ago = today_ts - pd.DateOffset(months=1)
                    jan_1     = pd.Timestamp("{}-01-01".format(date.today().year))

                    perf_rows = []
                    for fonds_n in pivot.columns:
                        series_n = pivot[fonds_n].dropna()
                        if series_n.empty:
                            continue
                        nav_last  = float(series_n.iloc[-1])
                        nav_first = float(series_n.iloc[0])

                        s1m  = series_n[series_n.index >= one_m_ago]
                        p1m  = ((nav_last / float(s1m.iloc[0]) - 1) * 100
                                if len(s1m) > 0 and float(s1m.iloc[0]) != 0 else float("nan"))

                        sytd = series_n[series_n.index >= jan_1]
                        pytd = ((nav_last / float(sytd.iloc[0]) - 1) * 100
                                if len(sytd) > 0 and float(sytd.iloc[0]) != 0 else float("nan"))

                        pp   = ((nav_last / nav_first - 1) * 100
                                if nav_first != 0 else float("nan"))

                        b100s = base100[fonds_n].dropna()
                        nb100 = float(b100s.iloc[-1]) if not b100s.empty else float("nan")

                        perf_rows.append({
                            "Fonds":            fonds_n,
                            "NAV Derniere":     round(nav_last, 4),
                            "Base 100 Actuel":  round(nb100, 2) if not np.isnan(nb100) else None,
                            "Perf 1M (%)":      round(p1m,  2) if not np.isnan(p1m)  else None,
                            "Perf YTD (%)":     round(pytd, 2) if not np.isnan(pytd) else None,
                            "Perf Periode (%)": round(pp,   2) if not np.isnan(pp)   else None,
                        })

                    if perf_rows:
                        df_pt = pd.DataFrame(perf_rows)
                        st.session_state["perf_data"] = df_pt

                        st.markdown("#### Tableau des Performances")

                        # Tableau HTML plat — zero helper
                        tbl_rows = []
                        for r in perf_rows:
                            def _fmt_pct(val):
                                if val is None:
                                    return '<span style="color:#999;">n.d.</span>'
                                c_pct = "#1a7a3c" if val >= 0 else "#8b2020"
                                sign  = "+" if val > 0 else ""
                                return '<span style="color:{};font-weight:700;">{}{:.2f}%</span>'.format(
                                    c_pct, sign, val)

                            b100_d = "{:.2f}".format(r["Base 100 Actuel"]) if r["Base 100 Actuel"] is not None else "n.d."
                            tbl_rows.append(
                                "<tr style=\"background:{bg};border-bottom:1px solid {gris};\">"
                                "<td style=\"padding:6px 11px;font-weight:600;color:{marine};\">{fonds}</td>"
                                "<td style=\"padding:6px 11px;text-align:right;color:{marine};\">{nav:.4f}</td>"
                                "<td style=\"padding:6px 11px;text-align:right;color:{marine};\">{b100}</td>"
                                "<td style=\"padding:6px 11px;text-align:right;\">{p1m}</td>"
                                "<td style=\"padding:6px 11px;text-align:right;\">{pytd}</td>"
                                "<td style=\"padding:6px 11px;text-align:right;\">{pp}</td>"
                                "</tr>".format(
                                    bg="#f7fafd" if perf_rows.index(r) % 2 == 0 else BLANC,
                                    gris=GRIS, marine=MARINE,
                                    fonds=r["Fonds"],
                                    nav=r["NAV Derniere"],
                                    b100=b100_d,
                                    p1m=_fmt_pct(r["Perf 1M (%)"]),
                                    pytd=_fmt_pct(r["Perf YTD (%)"]),
                                    pp=_fmt_pct(r["Perf Periode (%)"]),
                                )
                            )

                        tbl_html = (
                            "<table style=\"width:100%;border-collapse:collapse;font-size:0.79rem;\">"
                            "<thead><tr style=\"background:{marine};\">"
                            "<th style=\"padding:7px 11px;text-align:left;color:white;\">Fonds</th>"
                            "<th style=\"padding:7px 11px;text-align:right;color:white;\">NAV</th>"
                            "<th style=\"padding:7px 11px;text-align:right;color:white;\">Base 100</th>"
                            "<th style=\"padding:7px 11px;text-align:right;color:white;\">Perf 1M</th>"
                            "<th style=\"padding:7px 11px;text-align:right;color:white;\">Perf YTD</th>"
                            "<th style=\"padding:7px 11px;text-align:right;color:white;\">Perf Periode</th>"
                            "</tr></thead><tbody>"
                            + "".join(tbl_rows) +
                            "</tbody></table>"
                        ).format(marine=MARINE)
                        st.markdown(tbl_html, unsafe_allow_html=True)
                        st.markdown("<br>", unsafe_allow_html=True)

                        st.download_button(
                            "Exporter les performances CSV",
                            data=df_pt.to_csv(index=False).encode("utf-8"),
                            file_name="performances_{}.csv".format(date.today().isoformat()),
                            mime="text/csv",
                        )

                        # Graphique YTD — Plotly EXPLICITE
                        ytd_data = [(r["Fonds"], r["Perf YTD (%)"]) for r in perf_rows
                                    if r["Perf YTD (%)"] is not None]
                        if ytd_data:
                            ytd_data.sort(key=lambda x: x[1], reverse=True)
                            fonds_y = [x[0] for x in ytd_data]
                            vals_y  = [x[1] for x in ytd_data]
                            bar_cols_y = [MARINE if v >= 0 else "#aaaaaa" for v in vals_y]
                            text_y     = [("{:+.2f}%".format(v) if v >= 0 else "{:.2f}%".format(v))
                                          for v in vals_y]

                            st.markdown("#### Comparaison Performances YTD")
                            fig_ytd = go.Figure(go.Bar(
                                x=fonds_y,
                                y=vals_y,
                                marker_color=bar_cols_y,
                                marker_line_color=BLANC,
                                marker_line_width=0.5,
                                hovertemplate="<b>%{x}</b><br>YTD : %{y:.2f}%<extra></extra>",
                                text=text_y,
                                textposition="outside",
                                textfont_size=9,
                                textfont_color=MARINE,
                            ))
                            fig_ytd.add_hline(y=0, line_color=MARINE, line_width=0.8)
                            fig_ytd.update_layout(
                                title_text="Performance YTD par Fonds (%)",
                                title_font_size=13,
                                title_font_color=MARINE,
                                height=320,
                                paper_bgcolor=BLANC,
                                plot_bgcolor=BLANC,
                                font_color=MARINE,
                                font_family="Segoe UI, Arial",
                                showlegend=False,
                                xaxis_tickangle=-15,
                                xaxis_tickfont_size=10,
                                xaxis_tickfont_color=MARINE,
                                xaxis_showgrid=False,
                                yaxis_showgrid=True,
                                yaxis_gridcolor=GRIS,
                                yaxis_tickfont_size=9,
                                yaxis_tickfont_color=MARINE,
                                yaxis_ticksuffix="%",
                                yaxis_zeroline=False,
                                margin_l=10, margin_r=10, margin_t=40, margin_b=10,
                            )
                            st.plotly_chart(fig_ytd, use_container_width=True,
                                            config={"displayModeBar": False})

        except Exception as e:
            st.error("Erreur de traitement du fichier NAV : {}".format(e))
            import traceback
            with st.expander("Details de l'erreur"):
                st.code(traceback.format_exc())
    else:
        st.markdown(
            '<div style="background:#001c4b04;border:2px dashed #001c4b1a;'
            'padding:44px;text-align:center;margin-top:12px;">'
            '<div style="font-size:0.92rem;font-weight:700;color:{marine};">'
            'Module Performance et NAV</div>'
            '<div style="color:#777;font-size:0.79rem;max-width:380px;'
            'margin:0 auto;line-height:1.65;">'
            'Chargez un fichier Excel ou CSV avec les colonnes'
            ' <code>Date</code>, <code>Fonds</code>, <code>NAV</code>'
            ' pour generer les courbes Base 100 et le tableau de performances.'
            '<br><br>Utilisez le bouton <b>Generer fichier NAV demo</b>'
            ' pour un exemple testable immediatement.'
            '</div></div>'.format(marine=MARINE),
            unsafe_allow_html=True
        )
