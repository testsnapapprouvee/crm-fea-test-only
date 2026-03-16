# =============================================================================
# database.py  —  CRM Asset Management  —  Amundi Edition  v9.3 Production
# Priorite : STABILITE ABSOLUE — zero helper complexe, SQL plat
# Charte : #001c4b Marine | #019ee1 Ciel | #f07d00 Orange
# v9.3 :
#   - Base 100% vierge au démarrage (purge données démo)
#   - get_market_fonds_breakdown() : croisement Marché × Fonds pour Sales Intel
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
KYC_STATUTS         = ["Validé", "En cours", "Bloqué"]
PRODUCT_INTERESTS   = [
    "Global Value", "International Fund", "Income Builder",
    "Resilient Equity", "Private Debt", "Active ETFs",
]
ROLES_CONTACT = [
    "CIO", "CFO", "Gérant de portefeuille", "Analyste",
    "Responsable investissements", "Directeur général", "Autre",
]

# ---------------------------------------------------------------------------
# FORMATAGE FINANCIER — règle unique, jamais de chiffre brut affiché
# ---------------------------------------------------------------------------

def format_finance(val):
    """
    >= 1 Md  =>  "X.X Md€"
    sinon    =>  "X.X M€"
    """
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

    # --- TABLE CLIENTS (base) ---
    c.execute("""
        CREATE TABLE IF NOT EXISTS clients (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            nom_client  TEXT NOT NULL UNIQUE,
            type_client TEXT NOT NULL CHECK(type_client IN ('IFA','Wholesale','Instit','Family Office')),
            region      TEXT NOT NULL CHECK(region IN ('GCC','EMEA','APAC','Nordics','Asia ex-Japan','North America','LatAm')),
            created_at  DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # --- MIGRATIONS CLIENTS — colonnes CRM avancé ---
    _safe_alter(c, "ALTER TABLE clients ADD COLUMN parent_id INTEGER REFERENCES clients(id) ON DELETE SET NULL")
    _safe_alter(c, "ALTER TABLE clients ADD COLUMN tier TEXT DEFAULT 'Tier 2' CHECK(tier IN ('Tier 1','Tier 2','Tier 3'))")
    _safe_alter(c, "ALTER TABLE clients ADD COLUMN kyc_status TEXT DEFAULT 'En cours' CHECK(kyc_status IN ('Validé','En cours','Bloqué'))")
    _safe_alter(c, "ALTER TABLE clients ADD COLUMN product_interests TEXT DEFAULT ''")

    # --- TABLE CONTACTS (annuaire CRM) ---
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

    # --- TABLE PIPELINE ---
    c.execute("""
        CREATE TABLE IF NOT EXISTS pipeline (
            id                   INTEGER PRIMARY KEY AUTOINCREMENT,
            client_id            INTEGER NOT NULL,
            fonds                TEXT NOT NULL CHECK(fonds IN ('Global Value','International Fund','Income Builder','Resilient Equity','Private Debt','Active ETFs')),
            statut               TEXT NOT NULL DEFAULT 'Prospect' CHECK(statut IN ('Prospect','Initial Pitch','Due Diligence','Soft Commit','Funded','Paused','Lost','Redeemed')),
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

    # --- TABLE SALES TEAM ---
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

    # --- TABLE ACTIVITES ---
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

    # --- TABLE AUDIT LOG ---
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
    conn.close()


def _safe_alter(c, sql):
    """Exécute un ALTER TABLE silencieusement si la colonne existe déjà."""
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
    """
    Retourne tous les clients avec leur maison mère (parent).
    Colonnes : id, nom_client, type_client, region, tier, kyc_status, product_interests,
               parent_id, parent_nom (nom de la maison mère ou '')
    """
    conn = get_connection()
    df = pd.read_sql_query(
        "SELECT c.id, c.nom_client, c.type_client, c.region,"
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
    """Retourne la liste des commerciaux avec leur marché."""
    conn = get_connection()
    try:
        df = pd.read_sql_query("SELECT id, nom, marche FROM sales_team ORDER BY nom", conn)
    except Exception:
        df = pd.DataFrame(columns=["id","nom","marche"])
    conn.close()
    return df


def add_sales_member(nom, marche):
    """Ajoute un commercial à la table sales_team."""
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


def get_sales_by_marche(marche):
    """Retourne les commerciaux d'un marché donné."""
    conn = get_connection()
    c = conn.cursor()
    c.execute("SELECT nom FROM sales_team WHERE marche = ? ORDER BY nom", (marche,))
    names = [r[0] for r in c.fetchall()]
    conn.close()
    return names


# ---------------------------------------------------------------------------
# CONTACTS — ANNUAIRE CRM
# ---------------------------------------------------------------------------

def get_contacts(client_id):
    """Retourne tous les contacts d'un client."""
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
    """Ajoute un contact à un client."""
    conn = get_connection()
    c = conn.cursor()
    # Si is_primary, on dépromote les autres
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
# LECTURES — PIPELINE (avec derniere activite en colonne)
# ---------------------------------------------------------------------------

def get_pipeline_with_clients(fonds_filter=None):
    """Pipeline complet sans colonne activite (leger, pour les tableaux)."""
    fonds_sql, fonds_params = _fonds_clause(fonds_filter, "p")
    conn = get_connection()
    query = (
        "SELECT p.id, p.client_id, c.nom_client, c.type_client, c.region,"
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
    """
    Pipeline enrichi avec la derniere activite de chaque client.
    Colonne supplementaire : derniere_activite (str, peut etre vide).
    """
    fonds_sql, fonds_params = _fonds_clause(fonds_filter, "p")
    conn = get_connection()
    query = (
        "SELECT p.id, p.client_id, c.nom_client, c.type_client, c.region,"
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


def get_market_fonds_breakdown(mode="pipeline"):
    """
    Croise pipeline × sales_team pour produire :
      colonnes : marche, fonds, aum
    mode='pipeline' → AUM Révisé des deals actifs (Prospect…Soft Commit)
    mode='funded'   → AUM Financé des deals Funded
    Utilisé pour le graphique empilé Marché × Fonds dans Sales Tracking.
    """
    conn = get_connection()
    if mode == "funded":
        statut_clause = "p.statut = 'Funded'"
        aum_col       = "p.funded_aum"
    else:
        statut_clause = "p.statut IN ('Prospect','Initial Pitch','Due Diligence','Soft Commit')"
        aum_col       = "p.revised_aum"
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

    c.execute("SELECT c2.nom_client, c2.type_client, c2.region, p.fonds, p.funded_aum, p.sales_owner FROM pipeline p JOIN clients c2 ON c2.id=p.client_id WHERE p.statut='Redeemed' AND p.funded_aum > 0 " + fonds_sql + " ORDER BY p.funded_aum DESC LIMIT 10", fonds_params)
    outflows = [{k: (float(v) if k == "funded_aum" else str(v)) for k, v in dict(r).items()} for r in c.fetchall()]

    c.execute("SELECT p.statut, COUNT(*) FROM pipeline p WHERE 1=1 " + fonds_sql + " GROUP BY p.statut", fonds_params)
    statut_repartition = {r[0]: int(r[1]) for r in c.fetchall()}

    c.execute("SELECT p.fonds, COALESCE(SUM(p.funded_aum),0) FROM pipeline p WHERE p.statut='Funded' " + fonds_sql + " GROUP BY p.fonds ORDER BY 2 DESC", fonds_params)
    aum_by_fonds = {r[0]: float(r[1]) for r in c.fetchall()}

    c.execute(
        "SELECT SUM(p.revised_aum * COALESCE(p.closing_probability,50)/100.0)"
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

def add_client(nom_client, type_client, region,
               parent_id=None, tier="Tier 2",
               kyc_status="En cours", product_interests=""):
    """
    Crée un client avec les nouvelles colonnes CRM avancées.
    product_interests : str séparé par virgules (ex: "Global Value,Private Debt")
    """
    conn = get_connection()
    c = conn.cursor()
    pid = int(parent_id) if parent_id else None
    interests_str = product_interests if isinstance(product_interests, str) else ",".join(product_interests)
    c.execute(
        "INSERT INTO clients (nom_client, type_client, region, parent_id, tier, kyc_status, product_interests)"
        " VALUES (?,?,?,?,?,?,?)",
        (nom_client.strip(), type_client, region, pid, tier, kyc_status, interests_str)
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
    # funded_aum auto : si Funded et funded_aum=0, utiliser revised_aum
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
