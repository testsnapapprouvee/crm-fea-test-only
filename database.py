# =============================================================================
# database.py  —  CRM Asset Management  —  Amundi Edition
# Priorite : STABILITE ABSOLUE — zero helper complexe, SQL plat
# Charte : #001c4b Marine | #019ee1 Ciel | #f07d00 Orange
# Nouveau : get_pipeline_with_last_activity() — colonne derniere_activite
#           get_pipeline_by_statut()          — drill-down statut cliquable
#           get_overdue_actions_full()        — ligne complete pour modal retard
# =============================================================================

import sqlite3
import io
import pandas as pd
import numpy as np
from datetime import date, datetime, timedelta
from typing import Optional
import os

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                       "crm_asset_management.db")

_AUM_COLS           = ["target_aum_initial", "revised_aum", "funded_aum"]
_TEXT_NULLABLE_COLS = ["raison_perte", "concurrent_choisi"]
_AUDIT_CHAMPS       = ["statut", "fonds", "target_aum_initial",
                       "revised_aum", "funded_aum", "raison_perte"]

FONDS_REFERENTIEL = [
    "Global Value", "International Fund", "Income Builder",
    "Resilient Equity", "Private Debt", "Active ETFs",
]
REGIONS_REFERENTIEL = [
    "GCC", "EMEA", "APAC", "Nordics",
    "Asia ex-Japan", "North America", "LatAm",
]


# ---------------------------------------------------------------------------
# CONNEXION
# ---------------------------------------------------------------------------

def get_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


# ---------------------------------------------------------------------------
# HELPER — clause IN() sans crash si liste vide
# ---------------------------------------------------------------------------

def _fonds_clause(fonds_filter, alias="p"):
    if not fonds_filter:
        return "", []
    placeholders = ",".join("?" * len(fonds_filter))
    return " AND {}.fonds IN ({})".format(alias, placeholders), list(fonds_filter)


# ---------------------------------------------------------------------------
# NETTOYAGE DE TYPES
# ---------------------------------------------------------------------------

def _clean_pipeline_df(df):
    if df.empty:
        return df.copy()
    df = df.copy()
    for col in _AUM_COLS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0).astype("float64")
    if "next_action_date" in df.columns:
        parsed = pd.to_datetime(df["next_action_date"], errors="coerce")
        df["next_action_date"] = [ts.date() if pd.notna(ts) else None for ts in parsed]
    for col in _TEXT_NULLABLE_COLS:
        if col in df.columns:
            df[col] = df[col].apply(
                lambda v: "" if (v is None or (isinstance(v, float) and np.isnan(v))) else str(v)
            )
    for col in ["id", "client_id"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype("int64")
    return df


# ---------------------------------------------------------------------------
# INIT DB
# ---------------------------------------------------------------------------

def init_db():
    conn = get_connection()
    c = conn.cursor()

    c.execute("""
        CREATE TABLE IF NOT EXISTS clients (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            nom_client  TEXT NOT NULL UNIQUE,
            type_client TEXT NOT NULL CHECK(type_client IN ('IFA','Wholesale','Instit','Family Office')),
            region      TEXT NOT NULL CHECK(region IN ('GCC','EMEA','APAC','Nordics','Asia ex-Japan','North America','LatAm')),
            created_at  DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS pipeline (
            id                 INTEGER PRIMARY KEY AUTOINCREMENT,
            client_id          INTEGER NOT NULL,
            fonds              TEXT NOT NULL CHECK(fonds IN ('Global Value','International Fund','Income Builder','Resilient Equity','Private Debt','Active ETFs')),
            statut             TEXT NOT NULL DEFAULT 'Prospect' CHECK(statut IN ('Prospect','Initial Pitch','Due Diligence','Soft Commit','Funded','Paused','Lost','Redeemed')),
            target_aum_initial REAL DEFAULT 0,
            revised_aum        REAL DEFAULT 0,
            funded_aum         REAL DEFAULT 0,
            raison_perte       TEXT,
            concurrent_choisi  TEXT,
            next_action_date   DATE,
            sales_owner        TEXT NOT NULL DEFAULT 'Non assigne',
            created_at         DATETIME DEFAULT CURRENT_TIMESTAMP,
            updated_at         DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (client_id) REFERENCES clients(id) ON DELETE CASCADE
        )
    """)
    try:
        c.execute("ALTER TABLE pipeline ADD COLUMN sales_owner TEXT NOT NULL DEFAULT 'Non assigne'")
        conn.commit()
    except sqlite3.OperationalError:
        pass

    c.execute("""
        CREATE TABLE IF NOT EXISTS activites (
            id               INTEGER PRIMARY KEY AUTOINCREMENT,
            client_id        INTEGER NOT NULL,
            date             DATE NOT NULL,
            notes            TEXT,
            type_interaction TEXT,
            created_at       DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (client_id) REFERENCES clients(id) ON DELETE CASCADE
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS audit_log (
            id                INTEGER PRIMARY KEY AUTOINCREMENT,
            pipeline_id       INTEGER NOT NULL,
            champ_modifie     TEXT NOT NULL,
            ancienne_valeur   TEXT,
            nouvelle_valeur   TEXT,
            modified_by       TEXT NOT NULL DEFAULT 'system',
            date_modification DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (pipeline_id) REFERENCES pipeline(id) ON DELETE CASCADE
        )
    """)
    conn.commit()

    c.execute("SELECT COUNT(*) FROM clients")
    if c.fetchone()[0] == 0:
        _insert_demo_data(conn)
    conn.close()


def _insert_demo_data(conn):
    c = conn.cursor()
    today = date.today()
    clients = [
        ("Al Rajhi Capital",              "IFA",           "GCC"),
        ("Emirates NBD AM",               "Wholesale",     "GCC"),
        ("Norges Bank Investment Mgmt",   "Instit",        "Nordics"),
        ("GIC Singapore",                 "Instit",        "APAC"),
        ("Rothschild & Co Family Office", "Family Office", "EMEA"),
        ("BlackRock APAC Division",       "Wholesale",     "Asia ex-Japan"),
        ("JP Morgan Asset Management",    "Instit",        "North America"),
        ("ADIA Abu Dhabi",                "Instit",        "GCC"),
        ("Lombard Odier FO Geneva",       "Family Office", "EMEA"),
        ("BTG Pactual Wealth",            "Wholesale",     "LatAm"),
    ]
    c.executemany("INSERT INTO clients (nom_client, type_client, region) VALUES (?,?,?)", clients)
    pipeline = [
        (1,  "Global Value",       "Funded",        50_000_000,  45_000_000, 42_000_000, None,     None,       (today + timedelta(days=15)).isoformat(), "Sophie Laurent"),
        (2,  "Income Builder",     "Soft Commit",   30_000_000,  28_000_000, 0,           None,     None,       (today - timedelta(days=3)).isoformat(),  "Marc Dupont"),
        (3,  "Private Debt",       "Due Diligence", 100_000_000, 100_000_000,0,           None,     None,       (today + timedelta(days=7)).isoformat(),  "Sophie Laurent"),
        (4,  "Resilient Equity",   "Funded",        75_000_000,  80_000_000, 78_000_000, None,     None,       (today + timedelta(days=30)).isoformat(), "Karim Belhadj"),
        (5,  "International Fund", "Initial Pitch", 20_000_000,  20_000_000, 0,           None,     None,       (today - timedelta(days=10)).isoformat(), "Marc Dupont"),
        (6,  "Active ETFs",        "Lost",          40_000_000,  35_000_000, 0,           "Pricing","Vanguard", (today + timedelta(days=60)).isoformat(), "Karim Belhadj"),
        (7,  "Global Value",       "Funded",        120_000_000, 115_000_000,110_000_000, None,     None,       (today + timedelta(days=20)).isoformat(), "Sophie Laurent"),
        (8,  "Private Debt",       "Soft Commit",   200_000_000, 180_000_000,0,           None,     None,       (today - timedelta(days=1)).isoformat(),  "Karim Belhadj"),
        (9,  "Income Builder",     "Paused",        15_000_000,  12_000_000, 0,           "Macro",  "Internal", (today + timedelta(days=90)).isoformat(), "Marc Dupont"),
        (10, "Resilient Equity",   "Due Diligence", 60_000_000,  60_000_000, 0,           None,     None,       (today + timedelta(days=5)).isoformat(),  "Sophie Laurent"),
    ]
    c.executemany(
        "INSERT INTO pipeline (client_id, fonds, statut, target_aum_initial, revised_aum, funded_aum,"
        " raison_perte, concurrent_choisi, next_action_date, sales_owner) VALUES (?,?,?,?,?,?,?,?,?,?)",
        pipeline
    )
    activites = [
        (1,  (today - timedelta(days=5)).isoformat(),  "Suivi post-investissement Q1",       "Call"),
        (2,  (today - timedelta(days=2)).isoformat(),  "Envoi term sheet revise",             "Email"),
        (3,  (today - timedelta(days=1)).isoformat(),  "Due Diligence equipe risk",           "Meeting"),
        (4,  (today - timedelta(days=8)).isoformat(),  "Confirmation investissement",         "Email"),
        (7,  (today - timedelta(days=3)).isoformat(),  "Review trimestrielle performance",    "Meeting"),
        (8,  (today - timedelta(days=1)).isoformat(),  "Negociation conditions LP — validees","Call"),
        (5,  (today - timedelta(days=6)).isoformat(),  "Presentation initiale du fonds",      "Meeting"),
        (10, (today - timedelta(days=4)).isoformat(),  "Envoi DDQ complete",                  "Email"),
    ]
    c.executemany(
        "INSERT INTO activites (client_id, date, notes, type_interaction) VALUES (?,?,?,?)",
        activites
    )
    conn.commit()


# ---------------------------------------------------------------------------
# LECTURES — CLIENTS
# ---------------------------------------------------------------------------

def get_all_clients():
    conn = get_connection()
    df = pd.read_sql_query("SELECT id, nom_client, type_client, region FROM clients ORDER BY nom_client", conn)
    conn.close()
    return df


def get_client_options():
    df = get_all_clients()
    return {str(n): int(i) for n, i in zip(df["nom_client"], df["id"])}


def get_sales_owners():
    conn = get_connection()
    c = conn.cursor()
    c.execute("SELECT DISTINCT sales_owner FROM pipeline WHERE sales_owner != 'Non assigne' ORDER BY sales_owner")
    owners = [r[0] for r in c.fetchall()]
    conn.close()
    return owners


# ---------------------------------------------------------------------------
# LECTURES — PIPELINE (avec derniere activite en colonne)
# ---------------------------------------------------------------------------

def get_pipeline_with_clients(fonds_filter=None):
    """Pipeline complet sans colonne activite (leger, pour les tableaux)."""
    fonds_sql, fonds_params = _fonds_clause(fonds_filter, "p")
    conn = get_connection()
    query = (
        "SELECT p.id, p.client_id, c.nom_client, c.type_client, c.region,"
        " p.fonds, p.statut, p.target_aum_initial, p.revised_aum, p.funded_aum,"
        " p.raison_perte, p.concurrent_choisi, p.next_action_date, p.sales_owner"
        " FROM pipeline p JOIN clients c ON c.id = p.client_id"
        " WHERE 1=1 " + fonds_sql +
        " ORDER BY p.funded_aum DESC, p.revised_aum DESC"
    )
    df = pd.read_sql_query(query, conn, params=fonds_params if fonds_params else None)
    conn.close()
    return _clean_pipeline_df(df)


def get_pipeline_with_last_activity(fonds_filter=None):
    """
    Pipeline enrichi avec la derniere activite de chaque client.
    Colonne supplementaire : derniere_activite (str, peut etre vide).
    Utilise une sous-requete correlee SQLite — zero jointure cartesienne.
    """
    fonds_sql, fonds_params = _fonds_clause(fonds_filter, "p")
    conn = get_connection()
    query = (
        "SELECT p.id, p.client_id, c.nom_client, c.type_client, c.region,"
        " p.fonds, p.statut, p.target_aum_initial, p.revised_aum, p.funded_aum,"
        " p.raison_perte, p.concurrent_choisi, p.next_action_date, p.sales_owner,"
        " ("
        "   SELECT '[' || a.type_interaction || '] ' || a.notes"
        "   FROM activites a"
        "   WHERE a.client_id = p.client_id"
        "   ORDER BY a.date DESC, a.id DESC LIMIT 1"
        " ) AS derniere_activite"
        " FROM pipeline p JOIN clients c ON c.id = p.client_id"
        " WHERE 1=1 " + fonds_sql +
        " ORDER BY p.funded_aum DESC, p.revised_aum DESC"
    )
    df = pd.read_sql_query(query, conn, params=fonds_params if fonds_params else None)
    conn.close()
    df = _clean_pipeline_df(df)
    if "derniere_activite" in df.columns:
        df["derniere_activite"] = df["derniere_activite"].apply(
            lambda v: "" if (v is None or (isinstance(v, float) and np.isnan(v))) else str(v)
        )
    return df


def get_pipeline_by_statut(statut, fonds_filter=None):
    """
    Retourne tous les deals d'un statut donne, enrichis de la derniere activite.
    Utilise pour les modals drill-down des pastilles de statut.
    """
    fonds_sql, fonds_params = _fonds_clause(fonds_filter, "p")
    conn = get_connection()
    query = (
        "SELECT p.id, c.nom_client, c.type_client, c.region,"
        " p.fonds, p.statut,"
        " p.target_aum_initial, p.revised_aum, p.funded_aum,"
        " p.raison_perte, p.concurrent_choisi,"
        " p.next_action_date, p.sales_owner,"
        " ("
        "   SELECT '[' || a.type_interaction || '] ' || a.notes"
        "   FROM activites a WHERE a.client_id = p.client_id"
        "   ORDER BY a.date DESC, a.id DESC LIMIT 1"
        " ) AS derniere_activite"
        " FROM pipeline p JOIN clients c ON c.id = p.client_id"
        " WHERE p.statut = ? " + fonds_sql +
        " ORDER BY p.revised_aum DESC, p.funded_aum DESC"
    )
    params = [statut] + fonds_params
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    df = _clean_pipeline_df(df)
    if "derniere_activite" in df.columns:
        df["derniere_activite"] = df["derniere_activite"].apply(
            lambda v: "" if (v is None or (isinstance(v, float) and np.isnan(v))) else str(v)
        )
    return df


def get_pipeline_row_by_id(pipeline_id):
    conn = get_connection()
    df = pd.read_sql_query(
        "SELECT p.id, p.client_id, c.nom_client, c.type_client, c.region,"
        " p.fonds, p.statut, p.target_aum_initial, p.revised_aum, p.funded_aum,"
        " p.raison_perte, p.concurrent_choisi, p.next_action_date, p.sales_owner"
        " FROM pipeline p JOIN clients c ON c.id = p.client_id WHERE p.id = ?",
        conn, params=(int(pipeline_id),)
    )
    conn.close()
    if df.empty:
        return None
    df = _clean_pipeline_df(df)
    row = df.iloc[0].to_dict()
    for col in _AUM_COLS:
        row[col] = float(row.get(col) or 0.0)
    row["id"]          = int(row["id"])
    row["client_id"]   = int(row["client_id"])
    row["sales_owner"] = str(row.get("sales_owner") or "Non assigne")
    for col in _TEXT_NULLABLE_COLS:
        row[col] = str(row.get(col) or "")
    nad = row.get("next_action_date")
    if not isinstance(nad, date):
        row["next_action_date"] = date.today() + timedelta(days=14)
    return row


def get_overdue_actions(fonds_filter=None):
    """Liste legere des actions en retard (pour les alertes en-tete)."""
    today_str = date.today().isoformat()
    fonds_sql, fonds_params = _fonds_clause(fonds_filter, "p")
    conn = get_connection()
    query = (
        "SELECT p.id, c.nom_client, c.type_client, p.fonds, p.statut,"
        " p.next_action_date, p.sales_owner"
        " FROM pipeline p JOIN clients c ON c.id = p.client_id"
        " WHERE p.next_action_date < ? AND p.statut NOT IN ('Lost','Redeemed','Funded')"
        " " + fonds_sql +
        " ORDER BY p.next_action_date ASC"
    )
    df = pd.read_sql_query(query, conn, params=[today_str] + fonds_params)
    conn.close()
    return _clean_pipeline_df(df)


def get_overdue_deal_full(pipeline_id):
    """
    Retourne la ligne complete d'un deal en retard pour le modal de detail.
    Inclut la derniere activite.
    """
    conn = get_connection()
    df = pd.read_sql_query(
        "SELECT p.id, c.nom_client, c.type_client, c.region,"
        " p.fonds, p.statut, p.target_aum_initial, p.revised_aum, p.funded_aum,"
        " p.raison_perte, p.concurrent_choisi, p.next_action_date, p.sales_owner,"
        " ("
        "   SELECT '[' || a.type_interaction || '] ' || a.notes"
        "   FROM activites a WHERE a.client_id = p.client_id"
        "   ORDER BY a.date DESC, a.id DESC LIMIT 1"
        " ) AS derniere_activite"
        " FROM pipeline p JOIN clients c ON c.id = p.client_id WHERE p.id = ?",
        conn, params=(int(pipeline_id),)
    )
    conn.close()
    if df.empty:
        return None
    df = _clean_pipeline_df(df)
    row = df.iloc[0].to_dict()
    for col in _AUM_COLS:
        row[col] = float(row.get(col) or 0.0)
    row["id"] = int(row["id"])
    row["derniere_activite"] = str(row.get("derniere_activite") or "")
    nad = row.get("next_action_date")
    if not isinstance(nad, date):
        row["next_action_date"] = None
    return row


# ---------------------------------------------------------------------------
# LECTURES — DRILL-DOWN modales (funded / actifs / lost)
# ---------------------------------------------------------------------------

def get_funded_deals_detail(fonds_filter=None):
    fonds_sql, fonds_params = _fonds_clause(fonds_filter, "p")
    conn = get_connection()
    query = (
        "SELECT c.nom_client AS Client, p.fonds AS Fonds, c.type_client AS Type,"
        " c.region AS Region, p.funded_aum AS AUM_Finance,"
        " p.target_aum_initial AS AUM_Cible, p.sales_owner AS Commercial,"
        " p.next_action_date AS Prochaine_Action"
        " FROM pipeline p JOIN clients c ON c.id = p.client_id"
        " WHERE p.statut = 'Funded'" + fonds_sql +
        " ORDER BY p.funded_aum DESC"
    )
    df = pd.read_sql_query(query, conn, params=fonds_params if fonds_params else None)
    conn.close()
    if "AUM_Finance" in df.columns:
        df["AUM_Finance"] = pd.to_numeric(df["AUM_Finance"], errors="coerce").fillna(0.0)
    if "AUM_Cible" in df.columns:
        df["AUM_Cible"] = pd.to_numeric(df["AUM_Cible"], errors="coerce").fillna(0.0)
    return df


def get_active_deals_detail(fonds_filter=None):
    fonds_sql, fonds_params = _fonds_clause(fonds_filter, "p")
    conn = get_connection()
    query = (
        "SELECT c.nom_client AS Client, p.fonds AS Fonds, p.statut AS Statut,"
        " c.type_client AS Type, c.region AS Region,"
        " p.revised_aum AS AUM_Revise, p.target_aum_initial AS AUM_Cible,"
        " p.next_action_date AS Prochaine_Action, p.sales_owner AS Commercial"
        " FROM pipeline p JOIN clients c ON c.id = p.client_id"
        " WHERE p.statut IN ('Prospect','Initial Pitch','Due Diligence','Soft Commit')"
        " " + fonds_sql +
        " ORDER BY p.revised_aum DESC"
    )
    df = pd.read_sql_query(query, conn, params=fonds_params if fonds_params else None)
    conn.close()
    for col in ["AUM_Revise", "AUM_Cible"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
    if "Prochaine_Action" in df.columns:
        parsed = pd.to_datetime(df["Prochaine_Action"], errors="coerce")
        df["Prochaine_Action"] = [ts.date() if pd.notna(ts) else None for ts in parsed]
    return df


def get_lost_deals_detail(fonds_filter=None):
    fonds_sql, fonds_params = _fonds_clause(fonds_filter, "p")
    conn = get_connection()
    query = (
        "SELECT c.nom_client AS Client, p.fonds AS Fonds, p.statut AS Statut,"
        " p.target_aum_initial AS AUM_Cible, p.raison_perte AS Raison,"
        " p.concurrent_choisi AS Concurrent, p.sales_owner AS Commercial"
        " FROM pipeline p JOIN clients c ON c.id = p.client_id"
        " WHERE p.statut IN ('Lost','Paused')" + fonds_sql +
        " ORDER BY p.target_aum_initial DESC"
    )
    df = pd.read_sql_query(query, conn, params=fonds_params if fonds_params else None)
    conn.close()
    if "AUM_Cible" in df.columns:
        df["AUM_Cible"] = pd.to_numeric(df["AUM_Cible"], errors="coerce").fillna(0.0)
    for col in ["Raison", "Concurrent"]:
        if col in df.columns:
            df[col] = df[col].fillna("—").apply(
                lambda v: "—" if str(v).strip() in ("", "nan") else str(v)
            )
    return df


# ---------------------------------------------------------------------------
# LECTURES — AUDIT
# ---------------------------------------------------------------------------

def get_audit_log(pipeline_id):
    conn = get_connection()
    df = pd.read_sql_query(
        "SELECT champ_modifie AS Champ, ancienne_valeur AS Avant,"
        " nouvelle_valeur AS Apres, modified_by AS Par, date_modification AS Date"
        " FROM audit_log WHERE pipeline_id = ? ORDER BY date_modification DESC",
        conn, params=(int(pipeline_id),)
    )
    conn.close()
    return df


# ---------------------------------------------------------------------------
# LECTURES — SALES
# ---------------------------------------------------------------------------

def get_sales_metrics(fonds_filter=None):
    fonds_sql, fonds_params = _fonds_clause(fonds_filter, "p")
    conn = get_connection()
    query = (
        "SELECT p.sales_owner AS Commercial,"
        " COUNT(*) AS Nb_Deals,"
        " SUM(CASE WHEN p.statut='Funded' THEN 1 ELSE 0 END) AS Funded,"
        " SUM(CASE WHEN p.statut='Lost'   THEN 1 ELSE 0 END) AS Perdus,"
        " SUM(CASE WHEN p.statut IN ('Prospect','Initial Pitch','Due Diligence','Soft Commit')"
        "     THEN 1 ELSE 0 END) AS Actifs,"
        " COALESCE(SUM(CASE WHEN p.statut='Funded' THEN p.funded_aum ELSE 0 END),0) AS AUM_Finance,"
        " COALESCE(SUM(CASE WHEN p.statut IN ('Prospect','Initial Pitch','Due Diligence','Soft Commit')"
        "     THEN p.revised_aum ELSE 0 END),0) AS Pipeline_Actif,"
        " SUM(CASE WHEN p.next_action_date < DATE('now')"
        "     AND p.statut NOT IN ('Lost','Redeemed','Funded') THEN 1 ELSE 0 END) AS Retards"
        " FROM pipeline p WHERE 1=1 " + fonds_sql +
        " GROUP BY p.sales_owner ORDER BY AUM_Finance DESC"
    )
    df = pd.read_sql_query(query, conn, params=fonds_params if fonds_params else None)
    conn.close()
    for col in ["AUM_Finance", "Pipeline_Actif"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
    for col in ["Nb_Deals", "Funded", "Perdus", "Actifs", "Retards"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)
    return df


def get_next_actions_by_sales(days_ahead=30, fonds_filter=None):
    today_str = date.today().isoformat()
    limit_str = (date.today() + timedelta(days=days_ahead)).isoformat()
    fonds_sql, fonds_params = _fonds_clause(fonds_filter, "p")
    conn = get_connection()
    query = (
        "SELECT c.nom_client, p.fonds, p.statut, p.next_action_date,"
        " p.sales_owner, p.revised_aum"
        " FROM pipeline p JOIN clients c ON c.id = p.client_id"
        " WHERE p.next_action_date BETWEEN ? AND ?"
        " AND p.statut NOT IN ('Lost','Redeemed','Funded')"
        " " + fonds_sql +
        " ORDER BY p.next_action_date ASC"
    )
    df = pd.read_sql_query(query, conn, params=[today_str, limit_str] + fonds_params)
    conn.close()
    return _clean_pipeline_df(df)


# ---------------------------------------------------------------------------
# LECTURES — KPIs
# ---------------------------------------------------------------------------

def get_kpis(fonds_filter=None):
    fonds_sql, fonds_params = _fonds_clause(fonds_filter, "p")
    conn = get_connection()
    c = conn.cursor()

    c.execute("SELECT COALESCE(SUM(p.funded_aum),0) FROM pipeline p WHERE p.statut='Funded' " + fonds_sql, fonds_params)
    total_funded = float(c.fetchone()[0])

    c.execute("SELECT COALESCE(SUM(p.revised_aum),0) FROM pipeline p WHERE p.statut IN ('Prospect','Initial Pitch','Due Diligence','Soft Commit') " + fonds_sql, fonds_params)
    pipeline_actif = float(c.fetchone()[0])

    c.execute("SELECT COUNT(*) FROM pipeline p WHERE p.statut='Funded' " + fonds_sql, fonds_params)
    nb_funded = int(c.fetchone()[0])

    c.execute("SELECT COUNT(*) FROM pipeline p WHERE p.statut='Lost' " + fonds_sql, fonds_params)
    nb_lost = int(c.fetchone()[0])

    c.execute("SELECT COUNT(*) FROM pipeline p WHERE p.statut='Paused' " + fonds_sql, fonds_params)
    nb_paused = int(c.fetchone()[0])

    c.execute("SELECT COUNT(*) FROM pipeline p WHERE p.statut IN ('Prospect','Initial Pitch','Due Diligence','Soft Commit') " + fonds_sql, fonds_params)
    nb_actifs = int(c.fetchone()[0])

    taux = nb_funded / (nb_funded + nb_lost) * 100 if (nb_funded + nb_lost) > 0 else 0.0

    c.execute("SELECT c2.type_client, COALESCE(SUM(p.funded_aum),0) FROM pipeline p JOIN clients c2 ON c2.id=p.client_id WHERE p.statut='Funded' " + fonds_sql + " GROUP BY c2.type_client", fonds_params)
    aum_by_type = {r[0]: float(r[1]) for r in c.fetchall()}

    c.execute("SELECT c2.nom_client, c2.type_client, c2.region, p.fonds, p.funded_aum, p.sales_owner FROM pipeline p JOIN clients c2 ON c2.id=p.client_id WHERE p.statut='Funded' AND p.funded_aum > 0 " + fonds_sql + " ORDER BY p.funded_aum DESC LIMIT 10", fonds_params)
    top_deals = [{k: (float(v) if k == "funded_aum" else str(v)) for k, v in dict(r).items()} for r in c.fetchall()]

    # Outflows = Redeemed (AUM qui sort)
    c.execute("SELECT c2.nom_client, c2.type_client, c2.region, p.fonds, p.funded_aum, p.sales_owner FROM pipeline p JOIN clients c2 ON c2.id=p.client_id WHERE p.statut='Redeemed' AND p.funded_aum > 0 " + fonds_sql + " ORDER BY p.funded_aum DESC LIMIT 10", fonds_params)
    outflows = [{k: (float(v) if k == "funded_aum" else str(v)) for k, v in dict(r).items()} for r in c.fetchall()]

    c.execute("SELECT p.statut, COUNT(*) FROM pipeline p WHERE 1=1 " + fonds_sql + " GROUP BY p.statut", fonds_params)
    statut_repartition = {r[0]: int(r[1]) for r in c.fetchall()}

    c.execute("SELECT p.fonds, COALESCE(SUM(p.funded_aum),0) FROM pipeline p WHERE p.statut='Funded' " + fonds_sql + " GROUP BY p.fonds ORDER BY 2 DESC", fonds_params)
    aum_by_fonds = {r[0]: float(r[1]) for r in c.fetchall()}

    conn.close()
    return {
        "total_funded":       total_funded,
        "pipeline_actif":     pipeline_actif,
        "taux_conversion":    round(taux, 1),
        "nb_deals_actifs":    nb_actifs,
        "nb_funded":          nb_funded,
        "nb_lost":            nb_lost,
        "nb_paused":          nb_paused,
        "aum_by_type":        aum_by_type,
        "top_deals":          top_deals,
        "outflows":           outflows,
        "statut_repartition": statut_repartition,
        "aum_by_fonds":       aum_by_fonds,
    }


def get_aum_by_region(fonds_filter=None):
    fonds_sql, fonds_params = _fonds_clause(fonds_filter, "p")
    conn = get_connection()
    c = conn.cursor()
    c.execute("SELECT c2.region, COALESCE(SUM(p.funded_aum),0) AS aum FROM pipeline p JOIN clients c2 ON c2.id=p.client_id WHERE p.statut='Funded' " + fonds_sql + " GROUP BY c2.region ORDER BY aum DESC", fonds_params)
    result = {r[0]: float(r[1]) for r in c.fetchall() if float(r[1]) > 0}
    conn.close()
    return result


# ---------------------------------------------------------------------------
# LECTURES — ACTIVITES
# ---------------------------------------------------------------------------

def get_activities(client_id=None):
    conn = get_connection()
    query = "SELECT a.id, c.nom_client, a.date, a.type_interaction, a.notes FROM activites a JOIN clients c ON c.id = a.client_id"
    params = None
    if client_id:
        query += " WHERE a.client_id = ?"
        params = (int(client_id),)
    query += " ORDER BY a.date DESC LIMIT 50"
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df


# ---------------------------------------------------------------------------
# ECRITURE
# ---------------------------------------------------------------------------

def add_client(nom_client, type_client, region):
    conn = get_connection()
    c = conn.cursor()
    c.execute("INSERT INTO clients (nom_client, type_client, region) VALUES (?,?,?)",
              (nom_client.strip(), type_client, region))
    conn.commit()
    new_id = c.lastrowid
    conn.close()
    return new_id


def add_pipeline_entry(client_id, fonds, statut, target_aum, revised_aum,
                       funded_aum, raison_perte, concurrent_choisi,
                       next_action_date, sales_owner="Non assigne"):
    conn = get_connection()
    c = conn.cursor()
    c.execute(
        "INSERT INTO pipeline (client_id, fonds, statut, target_aum_initial, revised_aum, funded_aum,"
        " raison_perte, concurrent_choisi, next_action_date, sales_owner) VALUES (?,?,?,?,?,?,?,?,?,?)",
        (int(client_id), fonds, statut,
         float(target_aum or 0), float(revised_aum or 0), float(funded_aum or 0),
         raison_perte.strip() or None if raison_perte else None,
         concurrent_choisi.strip() or None if concurrent_choisi else None,
         next_action_date, sales_owner.strip() or "Non assigne")
    )
    conn.commit()
    new_id = c.lastrowid
    conn.close()
    return new_id


def update_pipeline_row(row, modified_by="utilisateur"):
    statut       = str(row.get("statut", ""))
    raison_perte = str(row.get("raison_perte") or "").strip()
    if statut in ("Lost", "Paused") and not raison_perte:
        return False, "Raison obligatoire pour le statut {}.".format(statut)

    nad = row.get("next_action_date")
    if isinstance(nad, (date, datetime)):
        nad_str = nad.strftime("%Y-%m-%d")
    elif isinstance(nad, str) and nad.strip():
        nad_str = nad.strip()
    else:
        nad_str = None

    pid         = int(row["id"])
    new_fonds   = str(row["fonds"])
    new_target  = float(row.get("target_aum_initial") or 0)
    new_revised = float(row.get("revised_aum") or 0)
    new_funded  = float(row.get("funded_aum") or 0)
    new_raison  = raison_perte or None
    new_conc    = str(row.get("concurrent_choisi") or "").strip() or None
    new_sales   = str(row.get("sales_owner") or "Non assigne").strip()

    conn = get_connection()
    c = conn.cursor()
    c.execute("SELECT fonds, statut, target_aum_initial, revised_aum, funded_aum, raison_perte FROM pipeline WHERE id = ?", (pid,))
    old_row = c.fetchone()
    if old_row:
        old_vals = {
            "fonds":              str(old_row["fonds"] or ""),
            "statut":             str(old_row["statut"] or ""),
            "target_aum_initial": float(old_row["target_aum_initial"] or 0),
            "revised_aum":        float(old_row["revised_aum"] or 0),
            "funded_aum":         float(old_row["funded_aum"] or 0),
            "raison_perte":       str(old_row["raison_perte"] or ""),
        }
        new_vals = {
            "fonds":              new_fonds, "statut": statut,
            "target_aum_initial": new_target, "revised_aum": new_revised,
            "funded_aum":         new_funded, "raison_perte": str(new_raison or ""),
        }
        for champ in _AUDIT_CHAMPS:
            ov = old_vals.get(champ)
            nv = new_vals.get(champ)
            if champ in ("target_aum_initial", "revised_aum", "funded_aum"):
                changed = abs(float(ov or 0) - float(nv or 0)) > 0.01
            else:
                changed = str(ov) != str(nv)
            if changed:
                c.execute(
                    "INSERT INTO audit_log (pipeline_id, champ_modifie, ancienne_valeur, nouvelle_valeur, modified_by) VALUES (?,?,?,?,?)",
                    (pid, champ, str(ov), str(nv), modified_by)
                )
    c.execute(
        "UPDATE pipeline SET fonds=?, statut=?, target_aum_initial=?, revised_aum=?,"
        " funded_aum=?, raison_perte=?, concurrent_choisi=?, next_action_date=?,"
        " sales_owner=?, updated_at=CURRENT_TIMESTAMP WHERE id=?",
        (new_fonds, statut, new_target, new_revised, new_funded,
         new_raison, new_conc, nad_str, new_sales, pid)
    )
    conn.commit()
    conn.close()
    return True, None


def add_activity(client_id, date_str, notes, type_interaction):
    conn = get_connection()
    conn.execute(
        "INSERT INTO activites (client_id, date, notes, type_interaction) VALUES (?,?,?,?)",
        (int(client_id), date_str, notes, type_interaction)
    )
    conn.commit()
    conn.close()


def upsert_clients_from_df(df):
    inserted, updated = 0, 0
    conn = get_connection()
    c = conn.cursor()
    df.columns = [col.lower().strip() for col in df.columns]
    for _, row in df.iterrows():
        nom = str(row.get("nom_client", "")).strip()
        if not nom:
            continue
        c.execute("SELECT id FROM clients WHERE nom_client=?", (nom,))
        if c.fetchone():
            c.execute("UPDATE clients SET type_client=?, region=? WHERE nom_client=?",
                      (str(row.get("type_client", "")), str(row.get("region", "")), nom))
            updated += 1
        else:
            c.execute("INSERT INTO clients (nom_client, type_client, region) VALUES (?,?,?)",
                      (nom, str(row.get("type_client", "")), str(row.get("region", ""))))
            inserted += 1
    conn.commit()
    conn.close()
    return inserted, updated


def upsert_pipeline_from_df(df):
    inserted, updated = 0, 0
    conn = get_connection()
    c = conn.cursor()
    df.columns = [col.lower().strip() for col in df.columns]
    for _, row in df.iterrows():
        client_id = None
        if "client_id" in df.columns and pd.notna(row.get("client_id")):
            client_id = int(row["client_id"])
        elif "nom_client" in df.columns:
            nom = str(row.get("nom_client", "")).strip()
            c.execute("SELECT id FROM clients WHERE nom_client=?", (nom,))
            res = c.fetchone()
            if res:
                client_id = res["id"]
        if client_id is None:
            continue
        fonds = str(row.get("fonds", "")).strip()
        if not fonds:
            continue
        statut  = str(row.get("statut", "Prospect"))
        raison  = str(row.get("raison_perte", "") or "").strip() or None
        conc    = str(row.get("concurrent_choisi", "") or "").strip() or None
        next_a  = str(row.get("next_action_date", "") or "").strip() or None
        t_aum   = float(row.get("target_aum_initial", 0) or 0)
        r_aum   = float(row.get("revised_aum", 0) or 0)
        f_aum   = float(row.get("funded_aum", 0) or 0)
        sales   = str(row.get("sales_owner", "Non assigne") or "Non assigne")
        c.execute("SELECT id FROM pipeline WHERE client_id=? AND fonds=?", (client_id, fonds))
        existing = c.fetchone()
        if existing:
            c.execute(
                "UPDATE pipeline SET statut=?, target_aum_initial=?, revised_aum=?, funded_aum=?,"
                " raison_perte=?, concurrent_choisi=?, next_action_date=?, sales_owner=?,"
                " updated_at=CURRENT_TIMESTAMP WHERE id=?",
                (statut, t_aum, r_aum, f_aum, raison, conc, next_a, sales, existing["id"])
            )
            updated += 1
        else:
            c.execute(
                "INSERT INTO pipeline (client_id, fonds, statut, target_aum_initial, revised_aum,"
                " funded_aum, raison_perte, concurrent_choisi, next_action_date, sales_owner)"
                " VALUES (?,?,?,?,?,?,?,?,?,?)",
                (client_id, fonds, statut, t_aum, r_aum, f_aum, raison, conc, next_a, sales)
            )
            inserted += 1
    conn.commit()
    conn.close()
    return inserted, updated


# ---------------------------------------------------------------------------
# BACKUP EXCEL
# ---------------------------------------------------------------------------

def get_excel_backup(fonds_filter=None):
    buf = io.BytesIO()
    df_p = get_pipeline_with_last_activity(fonds_filter).copy()
    for col in ["target_aum_initial", "revised_aum", "funded_aum"]:
        if col in df_p.columns:
            df_p[col] = (df_p[col] / 1_000_000).round(2)
    rename_p = {
        "id": "ID", "nom_client": "Client", "type_client": "Type", "region": "Region",
        "fonds": "Fonds", "statut": "Statut",
        "target_aum_initial": "AUM_Cible_M_EUR", "revised_aum": "AUM_Revise_M_EUR",
        "funded_aum": "AUM_Finance_M_EUR", "raison_perte": "Raison",
        "concurrent_choisi": "Concurrent", "next_action_date": "Prochaine_Action",
        "sales_owner": "Commercial", "derniere_activite": "Derniere_Activite",
    }
    df_p = df_p.rename(columns={k: v for k, v in rename_p.items() if k in df_p.columns})
    if "Prochaine_Action" in df_p.columns:
        df_p["Prochaine_Action"] = df_p["Prochaine_Action"].apply(
            lambda d: d.isoformat() if isinstance(d, date) else str(d or "")
        )
    conn = get_connection()
    df_a = pd.read_sql_query(
        "SELECT a.id, a.pipeline_id, c.nom_client, p.fonds,"
        " a.champ_modifie, a.ancienne_valeur, a.nouvelle_valeur,"
        " a.modified_by, a.date_modification"
        " FROM audit_log a JOIN pipeline p ON p.id = a.pipeline_id"
        " JOIN clients c ON c.id = p.client_id"
        " ORDER BY a.date_modification DESC LIMIT 200",
        conn
    )
    conn.close()
    df_c = get_all_clients()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_p.to_excel(writer, sheet_name="Pipeline_Actif",      index=False)
        df_a.to_excel(writer, sheet_name="Historique_Audit",    index=False)
        df_c.to_excel(writer, sheet_name="Referentiel_Clients", index=False)
    return buf.getvalue()
