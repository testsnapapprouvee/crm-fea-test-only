# =============================================================================
# database.py  —  CRM Asset Management  v16.0
# Priorite : STABILITE ABSOLUE — zero helper complexe, SQL plat
# Charte : #001c4b Marine | #019ee1 Ciel | #f07d00 Orange
# v16.0 :
#   - init_db : schema relaxe — suppression des CHECK restrictifs + migration
#   - get_dynamic_filters : SELECT DISTINCT region/fonds/statut depuis la DB
#   - update_sales_member / delete_sales_member : CRUD complet equipe commerciale
#   - upsert_pipeline_from_df : zero fake data — region/type lus depuis Excel
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

TIERS_REFERENTIEL   = ["Tier 1", "Tier 2", "Tier 3"]
KYC_STATUTS         = ["Valide", "En cours", "Bloque"]
PRODUCT_INTERESTS   = [
    "Global Value", "International Fund", "Income Builder",
    "Resilient Equity", "Private Debt", "Active ETFs",
]
ROLES_CONTACT = [
    "CIO", "CFO", "Gerant de portefeuille", "Analyste",
    "Responsable investissements", "Directeur general", "Autre",
]

# ---------------------------------------------------------------------------
# FORMATAGE FINANCIER
# ---------------------------------------------------------------------------

def format_finance(val):
    try:
        v = float(val)
    except (TypeError, ValueError):
        return "—"
    if v == 0:
        return "0.0 M€"
    if v >= 1_000_000_000:
        return "{:.1f} Md€".format(v / 1_000_000_000)
    return "{:.1f} M€".format(v / 1_000_000)


# ---------------------------------------------------------------------------
# CONNEXION
# ---------------------------------------------------------------------------

def get_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


# ---------------------------------------------------------------------------
# HELPER — clause IN()
# ---------------------------------------------------------------------------

def _fonds_clause(fonds_filter, alias="p"):
    if not fonds_filter:
        return "", []
    placeholders = ",".join("?" * len(fonds_filter))
    return " AND {}.fonds IN ({})".format(alias, placeholders), list(fonds_filter)


# ---------------------------------------------------------------------------
# CORRECTIF 1A — _clean_pipeline_df
# Force AUM numeric conversion BEFORE the df.empty guard
# This prevents "Expected numeric dtype" crashes on empty pipelines
# ---------------------------------------------------------------------------

def _clean_pipeline_df(df):
    df = df.copy()
    # Always convert AUM columns first — even on empty DataFrames
    # Force .astype(str) BEFORE pd.to_numeric to handle datetime.time objects
    # from malformed Excel cells (fix: float() argument must be a real number, not datetime.time)
    for col in _AUM_COLS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col].astype(str), errors="coerce").fillna(0.0).astype("float64")
    if df.empty:
        return df
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

def _migrate_relaxed_schema(conn):
    """
    Migration one-shot : supprime les CHECK restrictifs sur clients et pipeline
    en recréant les tables si nécessaire (SQLite ne supporte pas DROP CONSTRAINT).
    La migration est idempotente — elle détecte si elle a déjà été appliquée.
    """
    c = conn.cursor()
    # Inspecter le schéma actuel de la table clients
    c.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='clients'")
    row = c.fetchone()
    if row and row[0] and ("CHECK(type_client IN" in row[0] or "CHECK(region IN" in row[0]):
        # Recréation de la table clients sans CHECK restrictifs
        c.executescript("""
            PRAGMA foreign_keys = OFF;
            CREATE TABLE IF NOT EXISTS clients_new (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                nom_client  TEXT NOT NULL UNIQUE,
                type_client TEXT NOT NULL DEFAULT '',
                region      TEXT NOT NULL DEFAULT '',
                created_at  DATETIME DEFAULT CURRENT_TIMESTAMP,
                parent_id   INTEGER REFERENCES clients_new(id) ON DELETE SET NULL,
                tier        TEXT DEFAULT 'Tier 2',
                kyc_status  TEXT DEFAULT 'En cours',
                product_interests TEXT DEFAULT '',
                country     TEXT DEFAULT ''
            );
            INSERT OR IGNORE INTO clients_new
                SELECT id, nom_client, type_client, region, created_at,
                       parent_id, tier, kyc_status, product_interests, country
                FROM clients;
            DROP TABLE clients;
            ALTER TABLE clients_new RENAME TO clients;
            PRAGMA foreign_keys = ON;
        """)
        conn.commit()

    # Inspecter le schéma actuel de la table pipeline
    c.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='pipeline'")
    row_p = c.fetchone()
    if row_p and row_p[0] and "CHECK(fonds IN" in row_p[0]:
        c.executescript("""
            PRAGMA foreign_keys = OFF;
            CREATE TABLE IF NOT EXISTS pipeline_new (
                id                   INTEGER PRIMARY KEY AUTOINCREMENT,
                client_id            INTEGER NOT NULL,
                fonds                TEXT NOT NULL DEFAULT '',
                statut               TEXT NOT NULL DEFAULT 'Prospect',
                target_aum_initial   REAL DEFAULT 0,
                revised_aum          REAL DEFAULT 0,
                funded_aum           REAL DEFAULT 0,
                raison_perte         TEXT,
                concurrent_choisi    TEXT,
                next_action_date     DATE,
                sales_owner          TEXT NOT NULL DEFAULT 'Non assigne',
                closing_probability  REAL DEFAULT 50,
                created_at           DATETIME DEFAULT CURRENT_TIMESTAMP,
                updated_at           DATETIME DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (client_id) REFERENCES clients(id) ON DELETE CASCADE
            );
            INSERT OR IGNORE INTO pipeline_new
                SELECT id, client_id, fonds, statut,
                       target_aum_initial, revised_aum, funded_aum,
                       raison_perte, concurrent_choisi, next_action_date,
                       sales_owner, closing_probability, created_at, updated_at
                FROM pipeline;
            DROP TABLE pipeline;
            ALTER TABLE pipeline_new RENAME TO pipeline;
            PRAGMA foreign_keys = ON;
        """)
        conn.commit()


def init_db():
    conn = get_connection()
    c = conn.cursor()

    # Schema relaxé : pas de CHECK restrictifs sur type_client / region / fonds / statut
    c.execute("""
        CREATE TABLE IF NOT EXISTS clients (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            nom_client  TEXT NOT NULL UNIQUE,
            type_client TEXT NOT NULL DEFAULT '',
            region      TEXT NOT NULL DEFAULT '',
            created_at  DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    """)

    _safe_alter(c, "ALTER TABLE clients ADD COLUMN parent_id INTEGER REFERENCES clients(id) ON DELETE SET NULL")
    _safe_alter(c, "ALTER TABLE clients ADD COLUMN tier TEXT DEFAULT 'Tier 2'")
    _safe_alter(c, "ALTER TABLE clients ADD COLUMN kyc_status TEXT DEFAULT 'En cours'")
    _safe_alter(c, "ALTER TABLE clients ADD COLUMN product_interests TEXT DEFAULT ''")
    _safe_alter(c, "ALTER TABLE clients ADD COLUMN country TEXT DEFAULT ''")

    c.execute("""
        CREATE TABLE IF NOT EXISTS contacts (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            client_id  INTEGER NOT NULL,
            prenom     TEXT NOT NULL DEFAULT '',
            nom        TEXT NOT NULL DEFAULT '',
            role       TEXT DEFAULT '',
            email      TEXT DEFAULT '',
            telephone  TEXT DEFAULT '',
            linkedin   TEXT DEFAULT '',
            is_primary INTEGER DEFAULT 0,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (client_id) REFERENCES clients(id) ON DELETE CASCADE
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS pipeline (
            id                   INTEGER PRIMARY KEY AUTOINCREMENT,
            client_id            INTEGER NOT NULL,
            fonds                TEXT NOT NULL DEFAULT '',
            statut               TEXT NOT NULL DEFAULT 'Prospect',
            target_aum_initial   REAL DEFAULT 0,
            revised_aum          REAL DEFAULT 0,
            funded_aum           REAL DEFAULT 0,
            raison_perte         TEXT,
            concurrent_choisi    TEXT,
            next_action_date     DATE,
            sales_owner          TEXT NOT NULL DEFAULT 'Non assigne',
            closing_probability  REAL DEFAULT 50,
            created_at           DATETIME DEFAULT CURRENT_TIMESTAMP,
            updated_at           DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (client_id) REFERENCES clients(id) ON DELETE CASCADE
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS sales_team (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            nom        TEXT NOT NULL UNIQUE,
            marche     TEXT NOT NULL DEFAULT 'Global'
        )
    """)

    _safe_alter(c, "ALTER TABLE pipeline ADD COLUMN sales_owner TEXT NOT NULL DEFAULT 'Non assigne'")
    _safe_alter(c, "ALTER TABLE pipeline ADD COLUMN closing_probability REAL DEFAULT 50")
    conn.commit()

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

    # Migration rétroactive : supprime les CHECK restrictifs sur les DB existantes
    try:
        _migrate_relaxed_schema(conn)
    except Exception:
        pass  # Ne jamais bloquer le démarrage

    conn.close()


def _safe_alter(c, sql):
    try:
        c.execute(sql)
    except sqlite3.OperationalError:
        pass


# ---------------------------------------------------------------------------
# LECTURES — CLIENTS
# ---------------------------------------------------------------------------

def get_all_clients():
    conn = get_connection()
    df = pd.read_sql_query(
        "SELECT id, nom_client, type_client, region,"
        " COALESCE(country, '') AS country,"
        " COALESCE(parent_id, 0) AS parent_id,"
        " COALESCE(tier, 'Tier 2') AS tier,"
        " COALESCE(kyc_status, 'En cours') AS kyc_status,"
        " COALESCE(product_interests, '') AS product_interests"
        " FROM clients ORDER BY nom_client",
        conn
    )
    conn.close()
    return df


def get_client_options():
    df = get_all_clients()
    return {str(n): int(i) for n, i in zip(df["nom_client"], df["id"])}


def get_client_hierarchy():
    conn = get_connection()
    df = pd.read_sql_query(
        "SELECT c.id, c.nom_client, c.type_client, c.region,"
        " COALESCE(c.country, '') AS country,"
        " COALESCE(c.tier, 'Tier 2') AS tier,"
        " COALESCE(c.kyc_status, 'En cours') AS kyc_status,"
        " COALESCE(c.product_interests, '') AS product_interests,"
        " COALESCE(c.parent_id, 0) AS parent_id,"
        " COALESCE(p.nom_client, '') AS parent_nom"
        " FROM clients c"
        " LEFT JOIN clients p ON p.id = c.parent_id"
        " ORDER BY c.nom_client",
        conn
    )
    conn.close()
    return df


def get_sales_owners():
    conn = get_connection()
    c = conn.cursor()
    c.execute("SELECT nom FROM sales_team ORDER BY nom")
    rows = c.fetchall()
    if rows:
        owners = [r[0] for r in rows]
    else:
        c.execute("SELECT DISTINCT sales_owner FROM pipeline WHERE sales_owner != 'Non assigne' ORDER BY sales_owner")
        owners = [r[0] for r in c.fetchall()]
    conn.close()
    return owners


def get_sales_team():
    conn = get_connection()
    try:
        df = pd.read_sql_query("SELECT id, nom, marche FROM sales_team ORDER BY nom", conn)
    except Exception:
        df = pd.DataFrame(columns=["id","nom","marche"])
    conn.close()
    return df


def add_sales_member(nom, marche):
    conn = get_connection()
    c = conn.cursor()
    try:
        c.execute("INSERT OR IGNORE INTO sales_team (nom, marche) VALUES (?,?)",
                  (nom.strip(), marche.strip()))
        conn.commit()
        success = c.rowcount > 0
    except Exception:
        success = False
    conn.close()
    return success


def update_sales_member(sales_id: int, nom: str, marche: str):
    """Met à jour le nom et le marché d'un commercial existant."""
    try:
        conn = get_connection()
        c = conn.cursor()
        c.execute(
            "UPDATE sales_team SET nom=?, marche=? WHERE id=?",
            (nom.strip(), marche.strip(), int(sales_id))
        )
        conn.commit()
        conn.close()
        return True, None
    except Exception as e:
        return False, str(e)


def delete_sales_member(sales_id: int):
    """Supprime un commercial. Met à jour le pipeline : sales_owner -> 'Non assigne'."""
    try:
        conn = get_connection()
        c = conn.cursor()
        # Récupérer le nom avant suppression pour anonymiser le pipeline
        c.execute("SELECT nom FROM sales_team WHERE id=?", (int(sales_id),))
        row = c.fetchone()
        if row:
            nom_to_delete = row["nom"]
            c.execute(
                "UPDATE pipeline SET sales_owner='Non assigne'"
                " WHERE sales_owner=?",
                (nom_to_delete,)
            )
        c.execute("DELETE FROM sales_team WHERE id=?", (int(sales_id),))
        conn.commit()
        conn.close()
        return True, None
    except Exception as e:
        return False, str(e)


def get_sales_by_marche(marche):
    conn = get_connection()
    c = conn.cursor()
    c.execute("SELECT nom FROM sales_team WHERE marche = ? ORDER BY nom", (marche,))
    names = [r[0] for r in c.fetchall()]
    conn.close()
    return names


# ---------------------------------------------------------------------------
# CONTACTS
# ---------------------------------------------------------------------------

def get_contacts(client_id):
    conn = get_connection()
    df = pd.read_sql_query(
        "SELECT id, client_id, prenom, nom, role, email, telephone, linkedin, is_primary"
        " FROM contacts WHERE client_id = ?"
        " ORDER BY is_primary DESC, nom ASC",
        conn, params=(int(client_id),)
    )
    conn.close()
    return df


def add_contact(client_id, prenom, nom, role="", email="", telephone="", linkedin="", is_primary=False):
    conn = get_connection()
    c = conn.cursor()
    if is_primary:
        c.execute("UPDATE contacts SET is_primary=0 WHERE client_id=?", (int(client_id),))
    c.execute(
        "INSERT INTO contacts (client_id, prenom, nom, role, email, telephone, linkedin, is_primary)"
        " VALUES (?,?,?,?,?,?,?,?)",
        (int(client_id), prenom.strip(), nom.strip(),
         role.strip(), email.strip(), telephone.strip(),
         linkedin.strip(), 1 if is_primary else 0)
    )
    conn.commit()
    new_id = c.lastrowid
    conn.close()
    return new_id


# ---------------------------------------------------------------------------
# LECTURES — PIPELINE
# ---------------------------------------------------------------------------

def get_pipeline_with_clients(fonds_filter=None):
    fonds_sql, fonds_params = _fonds_clause(fonds_filter, "p")
    conn = get_connection()
    query = (
        "SELECT p.id, p.client_id, c.nom_client, c.type_client, c.region,"
        " COALESCE(c.country, '') AS country,"
        " p.fonds, p.statut, p.target_aum_initial, p.revised_aum, p.funded_aum,"
        " p.raison_perte, p.concurrent_choisi, p.next_action_date, p.sales_owner,"
        " COALESCE(p.closing_probability, 50) AS closing_probability"
        " FROM pipeline p JOIN clients c ON c.id = p.client_id"
        " WHERE 1=1 " + fonds_sql +
        " ORDER BY p.funded_aum DESC, p.revised_aum DESC"
    )
    df = pd.read_sql_query(query, conn, params=fonds_params if fonds_params else None)
    conn.close()
    return _clean_pipeline_df(df)


def get_pipeline_with_last_activity(fonds_filter=None):
    fonds_sql, fonds_params = _fonds_clause(fonds_filter, "p")
    conn = get_connection()
    query = (
        "SELECT p.id, p.client_id, c.nom_client, c.type_client, c.region,"
        " COALESCE(c.country, '') AS country,"
        " p.fonds, p.statut, p.target_aum_initial, p.revised_aum, p.funded_aum,"
        " p.raison_perte, p.concurrent_choisi, p.next_action_date, p.sales_owner,"
        " COALESCE(p.closing_probability, 50) AS closing_probability,"
        " COALESCE(c.kyc_status, 'En cours') AS kyc_status,"
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
    fonds_sql, fonds_params = _fonds_clause(fonds_filter, "p")
    conn = get_connection()
    query = (
        "SELECT p.id, c.nom_client, c.type_client, c.region,"
        " COALESCE(c.country, '') AS country,"
        " p.fonds, p.statut,"
        " p.target_aum_initial, p.revised_aum, p.funded_aum,"
        " (CASE WHEN p.revised_aum > 0 THEN p.revised_aum ELSE p.target_aum_initial END) AS aum_pipeline,"
        " p.raison_perte, p.concurrent_choisi,"
        " p.next_action_date, p.sales_owner,"
        " ("
        "   SELECT '[' || a.type_interaction || '] ' || a.notes"
        "   FROM activites a WHERE a.client_id = p.client_id"
        "   ORDER BY a.date DESC, a.id DESC LIMIT 1"
        " ) AS derniere_activite"
        " FROM pipeline p JOIN clients c ON c.id = p.client_id"
        " WHERE p.statut = ? " + fonds_sql +
        " ORDER BY aum_pipeline DESC, p.funded_aum DESC"
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
        " p.raison_perte, p.concurrent_choisi, p.next_action_date, p.sales_owner,"
        " COALESCE(p.closing_probability, 50) AS closing_probability,"
        " COALESCE(c.kyc_status, 'En cours') AS kyc_status"
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
    row["kyc_status"]  = str(row.get("kyc_status") or "En cours")
    for col in _TEXT_NULLABLE_COLS:
        row[col] = str(row.get(col) or "")
    nad = row.get("next_action_date")
    if not isinstance(nad, date):
        row["next_action_date"] = date.today() + timedelta(days=14)
    return row


def get_overdue_actions(fonds_filter=None):
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
# LECTURES — DRILL-DOWN
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
    for col in ["AUM_Finance", "AUM_Cible"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
    return df


def get_active_deals_detail(fonds_filter=None):
    fonds_sql, fonds_params = _fonds_clause(fonds_filter, "p")
    conn = get_connection()
    query = (
        "SELECT c.nom_client AS Client, p.fonds AS Fonds, p.statut AS Statut,"
        " c.type_client AS Type, c.region AS Region,"
        " (CASE WHEN p.revised_aum > 0 THEN p.revised_aum ELSE p.target_aum_initial END) AS AUM_Pipeline,"
        " p.revised_aum AS AUM_Revise, p.target_aum_initial AS AUM_Cible,"
        " p.next_action_date AS Prochaine_Action, p.sales_owner AS Commercial"
        " FROM pipeline p JOIN clients c ON c.id = p.client_id"
        " WHERE p.statut IN ('Prospect','Initial Pitch','Due Diligence','Soft Commit')"
        " " + fonds_sql +
        " ORDER BY AUM_Pipeline DESC"
    )
    df = pd.read_sql_query(query, conn, params=fonds_params if fonds_params else None)
    conn.close()
    for col in ["AUM_Pipeline", "AUM_Revise", "AUM_Cible"]:
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
# AUDIT
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
# SALES
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
        "     THEN (CASE WHEN p.revised_aum > 0 THEN p.revised_aum ELSE p.target_aum_initial END) ELSE 0 END),0) AS Pipeline_Actif,"
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


def get_market_fonds_breakdown(mode="pipeline"):
    conn = get_connection()
    if mode == "funded":
        statut_clause = "p.statut = 'Funded'"
        aum_col       = "p.funded_aum"
    else:
        statut_clause = "p.statut IN ('Prospect','Initial Pitch','Due Diligence','Soft Commit')"
        aum_col       = "(CASE WHEN p.revised_aum > 0 THEN p.revised_aum ELSE p.target_aum_initial END)"
    query = (
        "SELECT COALESCE(st.marche, p.sales_owner) AS marche,"
        " p.fonds,"
        " COALESCE(SUM({aum}), 0) AS aum"
        " FROM pipeline p"
        " LEFT JOIN sales_team st ON st.nom = p.sales_owner"
        " WHERE {statut}"
        " GROUP BY COALESCE(st.marche, p.sales_owner), p.fonds"
        " ORDER BY marche, p.fonds"
    ).format(aum=aum_col, statut=statut_clause)
    df = pd.read_sql_query(query, conn)
    conn.close()
    df["aum"] = pd.to_numeric(df["aum"], errors="coerce").fillna(0.0)
    return df


# ---------------------------------------------------------------------------
# KPIs
# ---------------------------------------------------------------------------

def get_kpis(fonds_filter=None):
    fonds_sql, fonds_params = _fonds_clause(fonds_filter, "p")
    conn = get_connection()
    c = conn.cursor()

    c.execute("SELECT COALESCE(SUM(p.funded_aum),0) FROM pipeline p WHERE p.statut='Funded' " + fonds_sql, fonds_params)
    total_funded = float(c.fetchone()[0])

    c.execute("SELECT COALESCE(SUM(CASE WHEN p.revised_aum > 0 THEN p.revised_aum ELSE p.target_aum_initial END),0) FROM pipeline p WHERE p.statut IN ('Prospect','Initial Pitch','Due Diligence','Soft Commit') " + fonds_sql, fonds_params)
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

    c.execute("SELECT c2.nom_client, c2.type_client, c2.region, p.fonds, p.funded_aum, p.sales_owner FROM pipeline p JOIN clients c2 ON c2.id=p.client_id WHERE p.statut='Redeemed' AND p.funded_aum > 0 " + fonds_sql + " ORDER BY p.funded_aum DESC LIMIT 10", fonds_params)
    outflows = [{k: (float(v) if k == "funded_aum" else str(v)) for k, v in dict(r).items()} for r in c.fetchall()]

    c.execute("SELECT p.statut, COUNT(*) FROM pipeline p WHERE 1=1 " + fonds_sql + " GROUP BY p.statut", fonds_params)
    statut_repartition = {r[0]: int(r[1]) for r in c.fetchall()}

    c.execute("SELECT p.fonds, COALESCE(SUM(p.funded_aum),0) FROM pipeline p WHERE p.statut='Funded' " + fonds_sql + " GROUP BY p.fonds ORDER BY 2 DESC", fonds_params)
    aum_by_fonds = {r[0]: float(r[1]) for r in c.fetchall()}

    c.execute(
        "SELECT SUM((CASE WHEN p.revised_aum > 0 THEN p.revised_aum ELSE p.target_aum_initial END)"
        " * COALESCE(p.closing_probability,50)/100.0)"
        " FROM pipeline p"
        " WHERE p.statut IN ('Prospect','Initial Pitch','Due Diligence','Soft Commit')"
        + fonds_sql,
        fonds_params
    )
    wp_row = c.fetchone()
    weighted_pipeline = float(wp_row[0]) if (wp_row and wp_row[0] is not None) else 0.0

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
        "weighted_pipeline":  weighted_pipeline,
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
# CHANTIER 2 — FILTRES DYNAMIQUES
# ---------------------------------------------------------------------------

def get_dynamic_filters():
    """
    Renvoie les listes de valeurs réellement présentes dans la DB.
    Utilisé pour alimenter les filtres multiselect de l'interface.
    Retourne un dict avec les clés :
      - regions  : liste triée des régions non-vides des clients
      - fonds    : liste triée des fonds présents dans le pipeline
      - statuts  : liste triée des statuts présents dans le pipeline
    Chaque liste a un fallback sur les référentiels si la DB est vide.
    """
    conn = get_connection()
    c = conn.cursor()

    c.execute(
        "SELECT DISTINCT c.region FROM clients c"
        " WHERE c.region IS NOT NULL AND c.region != ''"
        " ORDER BY c.region"
    )
    regions_db = [r[0] for r in c.fetchall()]

    c.execute(
        "SELECT DISTINCT p.fonds FROM pipeline p"
        " WHERE p.fonds IS NOT NULL AND p.fonds != ''"
        " ORDER BY p.fonds"
    )
    fonds_db = [r[0] for r in c.fetchall()]

    c.execute(
        "SELECT DISTINCT p.statut FROM pipeline p"
        " WHERE p.statut IS NOT NULL AND p.statut != ''"
        " ORDER BY p.statut"
    )
    statuts_db = [r[0] for r in c.fetchall()]

    conn.close()

    # Fallback sur les référentiels si la DB est vide (premier lancement)
    return {
        "regions": regions_db if regions_db else list(REGIONS_REFERENTIEL),
        "fonds":   fonds_db   if fonds_db   else list(FONDS_REFERENTIEL),
        "statuts": statuts_db if statuts_db else [
            "Prospect", "Initial Pitch", "Due Diligence",
            "Soft Commit", "Funded", "Paused", "Lost", "Redeemed"
        ],
    }


# ---------------------------------------------------------------------------
# ACTIVITES
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
# ECRITURE — CLIENTS / PIPELINE
# ---------------------------------------------------------------------------

def add_client(nom_client, type_client, region, country="",
               parent_id=None, tier="Tier 2",
               kyc_status="En cours", product_interests=""):
    conn = get_connection()
    c = conn.cursor()
    pid = int(parent_id) if parent_id else None
    interests_str = product_interests if isinstance(product_interests, str) else ",".join(product_interests)
    c.execute(
        "INSERT INTO clients"
        " (nom_client, type_client, region, country, parent_id, tier, kyc_status, product_interests)"
        " VALUES (?,?,?,?,?,?,?,?)",
        (nom_client.strip(), type_client, region, country.strip(),
         pid, tier, kyc_status, interests_str)
    )
    conn.commit()
    new_id = c.lastrowid
    conn.close()
    return new_id


def add_pipeline_entry(client_id, fonds, statut, target_aum, revised_aum,
                       funded_aum, raison_perte, concurrent_choisi,
                       next_action_date, sales_owner="Non assigne",
                       closing_probability=50):
    conn = get_connection()
    c = conn.cursor()
    fa = float(funded_aum or 0)
    if statut == "Funded" and fa == 0:
        fa = float(revised_aum or 0)
    c.execute(
        "INSERT INTO pipeline (client_id, fonds, statut, target_aum_initial, revised_aum, funded_aum,"
        " raison_perte, concurrent_choisi, next_action_date, sales_owner, closing_probability)"
        " VALUES (?,?,?,?,?,?,?,?,?,?,?)",
        (int(client_id), fonds, statut,
         float(target_aum or 0), float(revised_aum or 0), fa,
         raison_perte.strip() or None if raison_perte else None,
         concurrent_choisi.strip() or None if concurrent_choisi else None,
         next_action_date, sales_owner.strip() or "Non assigne",
         float(closing_probability or 50))
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
    if statut == "Funded" and new_funded == 0:
        new_funded = new_revised

    new_prob = float(row.get("closing_probability") or 50)
    c.execute(
        "UPDATE pipeline SET fonds=?, statut=?, target_aum_initial=?, revised_aum=?,"
        " funded_aum=?, raison_perte=?, concurrent_choisi=?, next_action_date=?,"
        " sales_owner=?, closing_probability=?, updated_at=CURRENT_TIMESTAMP WHERE id=?",
        (new_fonds, statut, new_target, new_revised, new_funded,
         new_raison, new_conc, nad_str, new_sales, new_prob, pid)
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


# ---------------------------------------------------------------------------
# CRUD — CLIENTS
# ---------------------------------------------------------------------------

def update_client(client_id, nom_client, type_client, region, country="",
                  parent_id=None, tier="Tier 2", kyc_status="En cours",
                  product_interests=""):
    try:
        conn = get_connection()
        c = conn.cursor()
        pid = int(parent_id) if parent_id else None
        interests_str = (product_interests if isinstance(product_interests, str)
                         else ",".join(product_interests))
        c.execute(
            "UPDATE clients SET nom_client=?, type_client=?, region=?, country=?,"
            " parent_id=?, tier=?, kyc_status=?, product_interests=?"
            " WHERE id=?",
            (nom_client.strip(), type_client, region, country.strip(),
             pid, tier, kyc_status, interests_str, int(client_id))
        )
        conn.commit()
        conn.close()
        return True, None
    except Exception as e:
        return False, str(e)


# ---------------------------------------------------------------------------
# CRUD — ACTIVITIES
# ---------------------------------------------------------------------------

def update_activity(activity_id, date_str, notes, type_interaction):
    try:
        conn = get_connection()
        conn.execute(
            "UPDATE activites SET date=?, notes=?, type_interaction=? WHERE id=?",
            (date_str, notes, type_interaction, int(activity_id))
        )
        conn.commit()
        conn.close()
        return True, None
    except Exception as e:
        return False, str(e)


def delete_activity(activity_id):
    try:
        conn = get_connection()
        conn.execute("DELETE FROM activites WHERE id=?", (int(activity_id),))
        conn.commit()
        conn.close()
        return True, None
    except Exception as e:
        return False, str(e)


# ---------------------------------------------------------------------------
# CRUD — CONTACTS
# ---------------------------------------------------------------------------

def update_contact(contact_id, prenom, nom, role="", email="",
                   telephone="", linkedin="", is_primary=False):
    try:
        conn = get_connection()
        c = conn.cursor()
        c.execute("SELECT client_id FROM contacts WHERE id=?", (int(contact_id),))
        row = c.fetchone()
        if not row:
            conn.close()
            return False, "Contact not found."
        client_id = row["client_id"]
        if is_primary:
            c.execute("UPDATE contacts SET is_primary=0 WHERE client_id=?", (client_id,))
        c.execute(
            "UPDATE contacts SET prenom=?, nom=?, role=?, email=?,"
            " telephone=?, linkedin=?, is_primary=? WHERE id=?",
            (prenom.strip(), nom.strip(), role.strip(), email.strip(),
             telephone.strip(), linkedin.strip(),
             1 if is_primary else 0, int(contact_id))
        )
        conn.commit()
        conn.close()
        return True, None
    except Exception as e:
        return False, str(e)


def delete_contact(contact_id):
    try:
        conn = get_connection()
        conn.execute("DELETE FROM contacts WHERE id=?", (int(contact_id),))
        conn.commit()
        conn.close()
        return True, None
    except Exception as e:
        return False, str(e)


# ---------------------------------------------------------------------------
# CRUD — PIPELINE
# ---------------------------------------------------------------------------

def delete_pipeline_row(pipeline_id):
    try:
        conn = get_connection()
        c = conn.cursor()
        c.execute("DELETE FROM audit_log WHERE pipeline_id=?", (int(pipeline_id),))
        c.execute("DELETE FROM pipeline WHERE id=?", (int(pipeline_id),))
        conn.commit()
        conn.close()
        return True, None
    except Exception as e:
        return False, str(e)


# ---------------------------------------------------------------------------
# MAILING LIST
# ---------------------------------------------------------------------------

def get_mailing_list(regions=None, countries=None, tiers=None, product_interests=None):
    conn = get_connection()
    query = (
        "SELECT ct.id AS contact_id, ct.prenom AS first_name, ct.nom AS last_name,"
        " c.nom_client AS company, ct.role, ct.email,"
        " c.region, COALESCE(c.country,'') AS country,"
        " COALESCE(c.tier,'Tier 2') AS tier,"
        " COALESCE(c.product_interests,'') AS product_interests"
        " FROM contacts ct"
        " JOIN clients c ON c.id = ct.client_id"
        " WHERE ct.email != '' AND ct.email IS NOT NULL"
        " ORDER BY c.nom_client, ct.nom"
    )
    df = pd.read_sql_query(query, conn)
    conn.close()
    if regions:
        df = df[df["region"].isin(regions)]
    if countries:
        df = df[df["country"].isin(countries)]
    if tiers:
        df = df[df["tier"].isin(tiers)]
    if product_interests:
        def _has_interest(pi_str):
            return any(p.strip() in pi_str for p in product_interests)
        df = df[df["product_interests"].apply(_has_interest)]
    return df.reset_index(drop=True)


def detect_import_duplicates(df):
    """Detect exact and fuzzy duplicate client names in an import DataFrame."""
    df = df.copy()
    df.columns = [col.lower().strip() for col in df.columns]
    col = "nom_client" if "nom_client" in df.columns else (
        "client" if "client" in df.columns else (
        "company" if "company" in df.columns else None))
    if col is None:
        return {"exact": [], "fuzzy": []}
    names = df[col].dropna().astype(str).str.strip().tolist()
    exact = []
    seen = {}
    for idx, n in enumerate(names, start=1):
        key = n.lower()
        if key in seen:
            exact.append({"ligne_1": seen[key], "ligne_2": idx, "valeur": n})
        else:
            seen[key] = idx
    fuzzy = []
    unique_names = list(dict.fromkeys(n.strip() for n in names if n.strip()))
    for i in range(len(unique_names)):
        for j in range(i + 1, len(unique_names)):
            a, b = unique_names[i].lower(), unique_names[j].lower()
            if a != b and (a in b or b in a or _simple_ratio(a, b) > 0.85):
                fuzzy.append({"nom_1": unique_names[i], "nom_2": unique_names[j]})
    return {"exact": exact, "fuzzy": fuzzy}


def _simple_ratio(a, b):
    """Simple similarity ratio based on common characters."""
    if not a or not b:
        return 0.0
    common = sum(1 for c in a if c in b)
    return 2.0 * common / (len(a) + len(b))


def delete_pipeline_rows(pipeline_ids):
    """Delete multiple pipeline rows and their audit logs in one transaction."""
    if not pipeline_ids:
        return 0
    try:
        conn = get_connection()
        c = conn.cursor()
        placeholders = ",".join("?" * len(pipeline_ids))
        ids = [int(pid) for pid in pipeline_ids]
        c.execute("DELETE FROM audit_log WHERE pipeline_id IN ({})".format(placeholders), ids)
        c.execute("DELETE FROM pipeline WHERE id IN ({})".format(placeholders), ids)
        deleted = c.rowcount
        conn.commit()
        conn.close()
        return deleted
    except Exception:
        return 0


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


# ---------------------------------------------------------------------------
# SMART IMPORT — Auto-cree clients et sales owners inconnus
# Regle metier : ZERO donnee inventee — region/type lus depuis Excel ou laisses vides
# ---------------------------------------------------------------------------

def upsert_pipeline_from_df(df):
    """
    Smart Import 100% autonome :
      A. nom_client inconnu  -> INSERT INTO clients (type_client et region lus depuis
         le fichier Excel, ou '' si absent — on n'invente jamais de valeur par defaut)
      B. sales_owner inconnu -> INSERT OR IGNORE INTO sales_team (marche='Global')
      Saute une ligne UNIQUEMENT si nom_client est totalement vide.
    Smart Mapper : standardise les noms de colonnes Excel entrants avant lecture.
    """
    # ── SMART MAPPER — standardisation des colonnes Excel entrantes ───────────
    # Convertit les variantes courantes vers les noms canoniques internes.
    _COLUMN_MAP = {
        # Client / Company
        "client":               "nom_client",
        "company":              "nom_client",
        "compte":               "nom_client",
        "nom client":           "nom_client",
        "nom_du_client":        "nom_client",
        # Type client
        "type":                 "type_client",
        "type client":          "type_client",
        "client type":          "type_client",
        "categorie":            "type_client",
        # Region
        "region":               "region",
        "geography":            "region",
        "geo":                  "region",
        "marche":               "region",
        # AUM Cible
        "aum cible":            "target_aum_initial",
        "aum_cible":            "target_aum_initial",
        "aum target":           "target_aum_initial",
        "target aum":           "target_aum_initial",
        "target_aum":           "target_aum_initial",
        "aum cible (m\u20ac)":  "target_aum_initial",
        "aum_cible_m_eur":      "target_aum_initial",
        # AUM Révisé
        "aum revise":           "revised_aum",
        "aum r\u00e9vis\u00e9": "revised_aum",
        "aum_revise":           "revised_aum",
        "aum revised":          "revised_aum",
        "revised aum":          "revised_aum",
        "aum revise (m\u20ac)": "revised_aum",
        "aum_revise_m_eur":     "revised_aum",
        # AUM Financé
        "aum finance":          "funded_aum",
        "aum financ\u00e9":     "funded_aum",
        "aum_finance":          "funded_aum",
        "funded aum":           "funded_aum",
        "aum finance (m\u20ac)":"funded_aum",
        "aum_finance_m_eur":    "funded_aum",
        # Commercial / Sales
        "commercial":           "sales_owner",
        "sales":                "sales_owner",
        "sales owner":          "sales_owner",
        "sales rep":            "sales_owner",
        "responsable":          "sales_owner",
        "charg\u00e9 d'affaires":"sales_owner",
        # Fonds
        "fund":                 "fonds",
        "produit":              "fonds",
        # Statut
        "status":               "statut",
        "stage":                "statut",
        "etape":                "statut",
        "\u00e9tape":           "statut",
        # Next Action
        "prochaine action":     "next_action_date",
        "prochaine_action":     "next_action_date",
        "next action":          "next_action_date",
        "date action":          "next_action_date",
        "date_action":          "next_action_date",
        # Probabilité
        "probabilite":          "closing_probability",
        "probabilit\u00e9":     "closing_probability",
        "proba":                "closing_probability",
        "probability":          "closing_probability",
        "closing probability":  "closing_probability",
        # Raison perte
        "raison":               "raison_perte",
        "raison perte":         "raison_perte",
        "loss reason":          "raison_perte",
        # Concurrent
        "concurrent":           "concurrent_choisi",
        "competitor":           "concurrent_choisi",
    }

    inserted, updated = 0, 0
    conn = get_connection()
    c = conn.cursor()
    # Normaliser d'abord en minuscules, puis appliquer le mapping
    df = df.copy()
    df.columns = [col.lower().strip() for col in df.columns]
    df.rename(columns=_COLUMN_MAP, inplace=True)

    for _, row in df.iterrows():
        # ── 1. Resoudre client_id ─────────────────────────────────────────────
        client_id = None

        # Priorite : client_id explicite dans le fichier
        if "client_id" in df.columns and pd.notna(row.get("client_id")):
            try:
                client_id = int(row["client_id"])
            except (ValueError, TypeError):
                pass

        # Sinon : resolution par nom_client
        if client_id is None:
            nom = str(row.get("nom_client", "")).strip()
            if not nom:
                continue  # Seul cas ou on saute la ligne

            c.execute("SELECT id FROM clients WHERE nom_client=?", (nom,))
            res = c.fetchone()
            if res:
                client_id = res["id"]
            else:
                # AUTO-CREATE : client absent -> creer avec les infos disponibles
                # Regle metier : on ne forge JAMAIS de données. Si absent = chaine vide.
                _type_client = str(row.get("type_client", "") or "").strip()
                _region      = str(row.get("region", "") or "").strip()
                try:
                    c.execute(
                        "INSERT INTO clients"
                        " (nom_client, type_client, region, tier, kyc_status)"
                        " VALUES (?, ?, ?, 'Tier 2', 'En cours')",
                        (nom, _type_client, _region)
                    )
                    client_id = c.lastrowid
                except sqlite3.IntegrityError:
                    # Race condition : re-lire
                    c.execute("SELECT id FROM clients WHERE nom_client=?", (nom,))
                    res2 = c.fetchone()
                    if res2:
                        client_id = res2["id"]
                    else:
                        continue

        if client_id is None:
            continue

        # ── 2. Valider le fonds ───────────────────────────────────────────────
        fonds = str(row.get("fonds", "")).strip()
        if not fonds:
            continue

        # ── 3. Preparer les champs ────────────────────────────────────────────
        statut = str(row.get("statut", "Prospect")).strip()
        raison = str(row.get("raison_perte", "") or "").strip() or None
        conc   = str(row.get("concurrent_choisi", "") or "").strip() or None
        next_a = str(row.get("next_action_date", "") or "").strip() or None

        def _safe_float(val, default=0.0):
            """Convert any value (including datetime.time) to float safely."""
            try:
                return float(pd.to_numeric(str(val), errors="coerce") or default)
            except Exception:
                return float(default)

        t_aum  = _safe_float(row.get("target_aum_initial", 0))
        r_aum  = _safe_float(row.get("revised_aum", 0))
        f_aum  = _safe_float(row.get("funded_aum", 0))
        sales  = str(row.get("sales_owner", "Non assigne") or "Non assigne").strip()
        prob   = _safe_float(row.get("closing_probability", 50), default=50.0)

        # ── 4. Auto-creer le sales owner s'il est inconnu ────────────────────
        if sales and sales != "Non assigne":
            c.execute("SELECT id FROM sales_team WHERE nom=?", (sales,))
            if not c.fetchone():
                try:
                    c.execute(
                        "INSERT OR IGNORE INTO sales_team (nom, marche) VALUES (?, 'Global')",
                        (sales,)
                    )
                except Exception:
                    pass  # Ignore les erreurs de contrainte

        # ── 5. Upsert pipeline ────────────────────────────────────────────────
        c.execute(
            "SELECT id FROM pipeline WHERE client_id=? AND fonds=?",
            (client_id, fonds)
        )
        existing = c.fetchone()
        if existing:
            c.execute(
                "UPDATE pipeline"
                " SET statut=?, target_aum_initial=?, revised_aum=?, funded_aum=?,"
                " raison_perte=?, concurrent_choisi=?, next_action_date=?,"
                " sales_owner=?, closing_probability=?,"
                " updated_at=CURRENT_TIMESTAMP"
                " WHERE id=?",
                (statut, t_aum, r_aum, f_aum, raison, conc,
                 next_a, sales, prob, existing["id"])
            )
            updated += 1
        else:
            c.execute(
                "INSERT INTO pipeline"
                " (client_id, fonds, statut, target_aum_initial, revised_aum,"
                "  funded_aum, raison_perte, concurrent_choisi, next_action_date,"
                "  sales_owner, closing_probability)"
                " VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                (client_id, fonds, statut, t_aum, r_aum, f_aum,
                 raison, conc, next_a, sales, prob)
            )
            inserted += 1

    conn.commit()
    conn.close()
    return inserted, updated


# ---------------------------------------------------------------------------
# ANALYTICS — Money in Motion & Whitespace
# ---------------------------------------------------------------------------

def get_expected_cashflows():
    """
    Projected inflows by closing month (next_action_date).
    weighted AUM = aum_pipeline * closing_probability / 100
    Returns: mois (YYYY-MM), fonds, aum_pondere — next 6 months only.
    """
    conn = get_connection()
    today_str = date.today().isoformat()
    limit_str = (date.today().replace(day=1) + timedelta(days=190)).isoformat()
    query = (
        "SELECT"
        " strftime('%Y-%m', COALESCE(p.next_action_date, date('now','+90 days'))) AS mois,"
        " p.fonds,"
        " SUM((CASE WHEN p.revised_aum > 0 THEN p.revised_aum ELSE p.target_aum_initial END)"
        "     * COALESCE(p.closing_probability, 50) / 100.0) AS aum_pondere"
        " FROM pipeline p"
        " WHERE p.statut IN ('Prospect','Initial Pitch','Due Diligence','Soft Commit')"
        " AND strftime('%Y-%m', COALESCE(p.next_action_date, date('now','+90 days')))"
        "     BETWEEN strftime('%Y-%m', ?) AND strftime('%Y-%m', ?)"
        " GROUP BY mois, p.fonds"
        " ORDER BY mois, p.fonds"
    )
    df = pd.read_sql_query(query, conn, params=[today_str, limit_str])
    conn.close()
    df["aum_pondere"] = pd.to_numeric(df["aum_pondere"], errors="coerce").fillna(0.0)
    return df


def get_whitespace_matrix():
    """
    Whitespace cross-sell analysis: clients x fonds.
    NaN = opportunity (not invested), float = funded_aum.
    """
    conn = get_connection()
    query = (
        "SELECT c.nom_client, p.fonds,"
        " COALESCE(SUM(CASE WHEN p.statut='Funded' THEN p.funded_aum ELSE 0 END), 0) AS funded_aum"
        " FROM pipeline p"
        " JOIN clients c ON c.id = p.client_id"
        " GROUP BY c.nom_client, p.fonds"
        " ORDER BY c.nom_client, p.fonds"
    )
    df = pd.read_sql_query(query, conn)
    conn.close()
    if df.empty:
        return pd.DataFrame()
    df["funded_aum"] = pd.to_numeric(df["funded_aum"], errors="coerce").fillna(0.0)
    df["funded_aum"] = df["funded_aum"].replace(0.0, float("nan"))
    pivot = df.pivot_table(index="nom_client", columns="fonds",
                           values="funded_aum", aggfunc="sum")
    for f in FONDS_REFERENTIEL:
        if f not in pivot.columns:
            pivot[f] = float("nan")
    return pivot[FONDS_REFERENTIEL]


def get_client_group_summary(client_id: int):
    """Consolidated AUM for client + all direct subsidiaries."""
    conn = get_connection()
    c = conn.cursor()

    c.execute(
        "SELECT id, nom_client, tier FROM clients WHERE parent_id = ? ORDER BY nom_client",
        (client_id,)
    )
    filiales_rows = [
        {"id": r["id"], "nom": r["nom_client"], "tier": r["tier"]}
        for r in c.fetchall()
    ]
    all_ids = [client_id] + [f["id"] for f in filiales_rows]
    placeholders = ",".join("?" * len(all_ids))

    c.execute(
        "SELECT COALESCE(SUM(funded_aum), 0) FROM pipeline"
        " WHERE client_id = ? AND statut = 'Funded'",
        (client_id,)
    )
    aum_direct = float(c.fetchone()[0])

    for fil in filiales_rows:
        c.execute(
            "SELECT COALESCE(SUM(funded_aum), 0) FROM pipeline"
            " WHERE client_id = ? AND statut = 'Funded'",
            (fil["id"],)
        )
        fil["aum"] = float(c.fetchone()[0])

    c.execute(
        "SELECT COALESCE(SUM(funded_aum), 0) FROM pipeline"
        " WHERE client_id IN ({}) AND statut = 'Funded'".format(placeholders),
        all_ids
    )
    aum_consolide = float(c.fetchone()[0])

    c.execute(
        "SELECT p.fonds, p.statut, p.next_action_date, p.sales_owner,"
        " (CASE WHEN p.revised_aum > 0 THEN p.revised_aum ELSE p.target_aum_initial END) AS aum_pipeline"
        " FROM pipeline p"
        " WHERE p.client_id IN ({})"
        " AND p.statut IN ('Prospect','Initial Pitch','Due Diligence','Soft Commit')"
        " ORDER BY p.next_action_date ASC LIMIT 5".format(placeholders),
        all_ids
    )
    next_actions = [
        {
            "fonds": r["fonds"], "statut": r["statut"],
            "nad": str(r["next_action_date"] or "—"),
            "aum_pipeline": float(r["aum_pipeline"] or 0),
            "sales_owner": str(r["sales_owner"] or ""),
        }
        for r in c.fetchall()
    ]

    c.execute(
        "SELECT DISTINCT fonds FROM pipeline"
        " WHERE client_id IN ({}) AND statut = 'Funded'".format(placeholders),
        all_ids
    )
    fonds_investis = [r[0] for r in c.fetchall()]

    conn.close()
    return {
        "aum_consolide":  aum_consolide,
        "aum_direct":     aum_direct,
        "filiales":       filiales_rows,
        "next_actions":   next_actions,
        "fonds_investis": fonds_investis,
    }


# ---------------------------------------------------------------------------
# CORRECTIF 1B — BACKUP EXCEL
# Force pd.to_numeric avant la division pour eviter "Expected numeric dtype"
# ---------------------------------------------------------------------------

def get_excel_backup(fonds_filter=None):
    buf = io.BytesIO()
    df_p = get_pipeline_with_last_activity(fonds_filter).copy()
    # Force numeric conversion BEFORE dividing — prevents crash on empty/mixed pipeline
    for col in ["target_aum_initial", "revised_aum", "funded_aum"]:
        if col in df_p.columns:
            df_p[col] = pd.to_numeric(df_p[col], errors="coerce").fillna(0.0)
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
