# =============================================================================
# database.py — Couche d'accès aux données SQLite
# CRM Asset Management — Charte Amundi
# Refactoring Staff Engineer : typage strict, zéro crash Streamlit
# =============================================================================

import sqlite3
import pandas as pd
import numpy as np
from datetime import date, datetime, timedelta
from typing import Optional
import os

DB_PATH = os.path.join(os.path.dirname(__file__), "crm_asset_management.db")

_AUM_COLS          = ["target_aum_initial", "revised_aum", "funded_aum"]
_TEXT_NULLABLE_COLS = ["raison_perte", "concurrent_choisi"]


# ---------------------------------------------------------------------------
# CONNEXION
# ---------------------------------------------------------------------------

def get_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


# ---------------------------------------------------------------------------
# NETTOYAGE DE TYPES — FONCTION CENTRALE (anti-crash StreamlitAPIException)
# ---------------------------------------------------------------------------

def _clean_pipeline_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normalise TOUS les types du DataFrame pipeline avant exposition à Streamlit.

    Garanties :
    - AUM cols      → float64, NaN → 0.0
    - next_action_date → datetime.date Python natif | None  (jamais NaT, jamais str)
    - Texte nullable   → str Python "", jamais NaN / None
    - id, client_id    → int Python natif (évite numpy.int64)
    """
    if df.empty:
        return df.copy()

    df = df.copy()

    # 1. AUM : float64, fillna(0.0)
    for col in _AUM_COLS:
        if col in df.columns:
            df[col] = (
                pd.to_numeric(df[col], errors="coerce")
                .fillna(0.0)
                .astype("float64")
            )

    # 2. next_action_date → datetime.date | None  (jamais NaT ni str)
    if "next_action_date" in df.columns:
        parsed = pd.to_datetime(df["next_action_date"], errors="coerce")
        df["next_action_date"] = [
            ts.date() if pd.notna(ts) else None
            for ts in parsed
        ]

    # 3. Colonnes texte nullable → str, None/NaN → ""
    for col in _TEXT_NULLABLE_COLS:
        if col in df.columns:
            df[col] = df[col].apply(
                lambda v: ""
                if (v is None or (isinstance(v, float) and np.isnan(v)))
                else str(v)
            )

    # 4. Colonnes entières → int Python natif
    for col in ["id", "client_id"]:
        if col in df.columns:
            df[col] = (
                pd.to_numeric(df[col], errors="coerce")
                .fillna(0)
                .astype("int64")
            )

    return df


# ---------------------------------------------------------------------------
# INIT DB + DONNÉES FICTIVES
# ---------------------------------------------------------------------------

def init_db():
    """Crée les tables et insère des données fictives si la base est vide."""
    conn = get_connection()
    c = conn.cursor()

    c.execute("""
        CREATE TABLE IF NOT EXISTS clients (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            nom_client  TEXT NOT NULL UNIQUE,
            type_client TEXT NOT NULL
                        CHECK(type_client IN ('IFA','Wholesale','Instit','Family Office')),
            region      TEXT NOT NULL
                        CHECK(region IN ('GCC','EMEA','APAC','Nordics',
                                         'Asia ex-Japan','North America','LatAm')),
            created_at  DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS pipeline (
            id                 INTEGER PRIMARY KEY AUTOINCREMENT,
            client_id          INTEGER NOT NULL,
            fonds              TEXT NOT NULL
                               CHECK(fonds IN ('Global Value','International Fund',
                                               'Income Builder','Resilient Equity',
                                               'Private Debt','Active ETFs')),
            statut             TEXT NOT NULL DEFAULT 'Prospect'
                               CHECK(statut IN ('Prospect','Initial Pitch','Due Diligence',
                                                'Soft Commit','Funded','Paused','Lost','Redeemed')),
            target_aum_initial REAL DEFAULT 0,
            revised_aum        REAL DEFAULT 0,
            funded_aum         REAL DEFAULT 0,
            raison_perte       TEXT,
            concurrent_choisi  TEXT,
            next_action_date   DATE,
            created_at         DATETIME DEFAULT CURRENT_TIMESTAMP,
            updated_at         DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (client_id) REFERENCES clients(id) ON DELETE CASCADE
        )
    """)

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

    conn.commit()

    c.execute("SELECT COUNT(*) FROM clients")
    if c.fetchone()[0] == 0:
        _insert_dummy_data(conn)

    conn.close()


def _insert_dummy_data(conn: sqlite3.Connection):
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
    c.executemany(
        "INSERT INTO clients (nom_client, type_client, region) VALUES (?,?,?)",
        clients
    )

    pipeline = [
        (1,  "Global Value",       "Funded",        50_000_000,  45_000_000, 42_000_000, None, None,
             (today + timedelta(days=15)).isoformat()),
        (2,  "Income Builder",     "Soft Commit",   30_000_000,  28_000_000,          0, None, None,
             (today - timedelta(days=3)).isoformat()),
        (3,  "Private Debt",       "Due Diligence", 100_000_000, 100_000_000,         0, None, None,
             (today + timedelta(days=7)).isoformat()),
        (4,  "Resilient Equity",   "Funded",         75_000_000,  80_000_000, 78_000_000, None, None,
             (today + timedelta(days=30)).isoformat()),
        (5,  "International Fund", "Initial Pitch",  20_000_000,  20_000_000,          0, None, None,
             (today - timedelta(days=10)).isoformat()),
        (6,  "Active ETFs",        "Lost",           40_000_000,  35_000_000,          0, "Pricing", "Vanguard",
             (today + timedelta(days=60)).isoformat()),
        (7,  "Global Value",       "Funded",        120_000_000, 115_000_000, 110_000_000, None, None,
             (today + timedelta(days=20)).isoformat()),
        (8,  "Private Debt",       "Soft Commit",   200_000_000, 180_000_000,          0, None, None,
             (today - timedelta(days=1)).isoformat()),
        (9,  "Income Builder",     "Paused",         15_000_000,  12_000_000,          0, "Macro", "Internal",
             (today + timedelta(days=90)).isoformat()),
        (10, "Resilient Equity",   "Due Diligence",  60_000_000,  60_000_000,          0, None, None,
             (today + timedelta(days=5)).isoformat()),
    ]
    c.executemany("""
        INSERT INTO pipeline
            (client_id, fonds, statut, target_aum_initial, revised_aum, funded_aum,
             raison_perte, concurrent_choisi, next_action_date)
        VALUES (?,?,?,?,?,?,?,?,?)
    """, pipeline)

    activites = [
        (1,  (today - timedelta(days=5)).isoformat(),  "Call de suivi post-investissement Q1",        "Call"),
        (2,  (today - timedelta(days=2)).isoformat(),  "Envoi du term sheet révisé",                   "Email"),
        (3,  (today - timedelta(days=1)).isoformat(),  "Réunion Due Diligence avec l'équipe risk",     "Meeting"),
        (4,  (today - timedelta(days=8)).isoformat(),  "Confirmation d'investissement reçue",           "Email"),
        (7,  (today - timedelta(days=3)).isoformat(),  "Review trimestrielle de la performance fonds",  "Meeting"),
        (8,  (today - timedelta(days=1)).isoformat(),  "Call de négociation des conditions LP",         "Call"),
        (5,  (today - timedelta(days=6)).isoformat(),  "Présentation initiale du fonds International",  "Meeting"),
        (10, (today - timedelta(days=4)).isoformat(),  "Envoi DDQ complété",                            "Email"),
    ]
    c.executemany(
        "INSERT INTO activites (client_id, date, notes, type_interaction) VALUES (?,?,?,?)",
        activites
    )
    conn.commit()


# ---------------------------------------------------------------------------
# LECTURES — CLIENTS
# ---------------------------------------------------------------------------

def get_all_clients() -> pd.DataFrame:
    conn = get_connection()
    df = pd.read_sql_query(
        "SELECT id, nom_client, type_client, region FROM clients ORDER BY nom_client",
        conn
    )
    conn.close()
    return df


def get_client_options() -> dict:
    """Dict {nom_client: id (int Python)} pour les selectbox."""
    df = get_all_clients()
    return {str(nom): int(cid) for nom, cid in zip(df["nom_client"], df["id"])}


# ---------------------------------------------------------------------------
# LECTURES — PIPELINE
# ---------------------------------------------------------------------------

def get_pipeline_with_clients() -> pd.DataFrame:
    """
    Pipeline complet enrichi des infos client.
    Tous les types sont normalisés — zéro NaT, zéro numpy scalar, zéro str de date.
    """
    conn = get_connection()
    df = pd.read_sql_query("""
        SELECT
            p.id, p.client_id,
            c.nom_client, c.type_client, c.region,
            p.fonds, p.statut,
            p.target_aum_initial, p.revised_aum, p.funded_aum,
            p.raison_perte, p.concurrent_choisi, p.next_action_date
        FROM pipeline p
        JOIN clients c ON c.id = p.client_id
        ORDER BY p.funded_aum DESC, p.revised_aum DESC
    """, conn)
    conn.close()
    return _clean_pipeline_df(df)


def get_pipeline_row_by_id(pipeline_id: int) -> Optional[dict]:
    """
    Retourne une seule ligne pipeline sous forme de dict Python natif.
    Utilisé par le formulaire Master-Detail pour pré-remplir les champs.
    Garantit : AUM = float, next_action_date = datetime.date, textes = str.
    """
    conn = get_connection()
    df = pd.read_sql_query("""
        SELECT
            p.id, p.client_id, c.nom_client, c.type_client, c.region,
            p.fonds, p.statut, p.target_aum_initial, p.revised_aum, p.funded_aum,
            p.raison_perte, p.concurrent_choisi, p.next_action_date
        FROM pipeline p
        JOIN clients c ON c.id = p.client_id
        WHERE p.id = ?
    """, conn, params=(int(pipeline_id),))
    conn.close()

    if df.empty:
        return None

    df = _clean_pipeline_df(df)
    row = df.iloc[0].to_dict()

    # Garanties supplémentaires — types Python stricts
    for col in _AUM_COLS:
        row[col] = float(row.get(col) or 0.0)

    row["id"]        = int(row["id"])
    row["client_id"] = int(row["client_id"])

    for col in _TEXT_NULLABLE_COLS:
        row[col] = str(row.get(col) or "")

    # next_action_date → datetime.date (jamais None dans le form)
    nad = row.get("next_action_date")
    if not isinstance(nad, date):
        row["next_action_date"] = date.today() + timedelta(days=14)

    return row


def get_overdue_actions() -> pd.DataFrame:
    """Deals avec next_action_date dépassée (statut actif uniquement)."""
    today_str = date.today().isoformat()
    conn = get_connection()
    df = pd.read_sql_query(f"""
        SELECT c.nom_client, c.type_client, p.fonds, p.statut, p.next_action_date
        FROM pipeline p
        JOIN clients c ON c.id = p.client_id
        WHERE p.next_action_date < '{today_str}'
          AND p.statut NOT IN ('Lost','Redeemed','Funded')
        ORDER BY p.next_action_date ASC
    """, conn)
    conn.close()
    return _clean_pipeline_df(df)


# ---------------------------------------------------------------------------
# LECTURES — ACTIVITÉS
# ---------------------------------------------------------------------------

def get_activities(client_id: Optional[int] = None) -> pd.DataFrame:
    conn = get_connection()
    query = """
        SELECT a.id, c.nom_client, a.date, a.type_interaction, a.notes
        FROM activites a JOIN clients c ON c.id = a.client_id
    """
    params = None
    if client_id:
        query += " WHERE a.client_id = ?"
        params = (int(client_id),)
    query += " ORDER BY a.date DESC LIMIT 50"
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df


# ---------------------------------------------------------------------------
# ÉCRITURE — CLIENTS
# ---------------------------------------------------------------------------

def add_client(nom_client: str, type_client: str, region: str) -> int:
    conn = get_connection()
    c = conn.cursor()
    c.execute(
        "INSERT INTO clients (nom_client, type_client, region) VALUES (?,?,?)",
        (nom_client.strip(), type_client, region)
    )
    conn.commit()
    new_id = c.lastrowid
    conn.close()
    return new_id


def update_client(client_id: int, nom_client: str, type_client: str, region: str):
    conn = get_connection()
    conn.execute(
        "UPDATE clients SET nom_client=?, type_client=?, region=? WHERE id=?",
        (nom_client.strip(), type_client, region, int(client_id))
    )
    conn.commit()
    conn.close()


# ---------------------------------------------------------------------------
# ÉCRITURE — PIPELINE
# ---------------------------------------------------------------------------

def add_pipeline_entry(
    client_id: int, fonds: str, statut: str,
    target_aum: float, revised_aum: float, funded_aum: float,
    raison_perte: str, concurrent_choisi: str, next_action_date: str
) -> int:
    conn = get_connection()
    c = conn.cursor()
    c.execute("""
        INSERT INTO pipeline
            (client_id, fonds, statut, target_aum_initial, revised_aum, funded_aum,
             raison_perte, concurrent_choisi, next_action_date)
        VALUES (?,?,?,?,?,?,?,?,?)
    """, (
        int(client_id), fonds, statut,
        float(target_aum or 0), float(revised_aum or 0), float(funded_aum or 0),
        raison_perte.strip() or None if raison_perte else None,
        concurrent_choisi.strip() or None if concurrent_choisi else None,
        next_action_date
    ))
    conn.commit()
    new_id = c.lastrowid
    conn.close()
    return new_id


def update_pipeline_row(row: dict) -> tuple:
    """
    Met à jour une ligne pipeline.
    Validation : raison_perte obligatoire si statut Lost/Paused.
    Retourne (True, None) si OK, (False, message_erreur) sinon.
    """
    statut       = str(row.get("statut", ""))
    raison_perte = str(row.get("raison_perte") or "").strip()

    if statut in ("Lost", "Paused") and not raison_perte:
        return False, f"⚠️ La raison de perte/pause est obligatoire pour le statut « {statut} »."

    # Normalisation next_action_date → str ISO
    nad = row.get("next_action_date")
    if isinstance(nad, (date, datetime)):
        nad_str = nad.strftime("%Y-%m-%d")
    elif isinstance(nad, str) and nad.strip():
        nad_str = nad.strip()
    else:
        nad_str = None

    conn = get_connection()
    conn.execute("""
        UPDATE pipeline SET
            fonds=?, statut=?,
            target_aum_initial=?, revised_aum=?, funded_aum=?,
            raison_perte=?, concurrent_choisi=?,
            next_action_date=?,
            updated_at=CURRENT_TIMESTAMP
        WHERE id=?
    """, (
        str(row["fonds"]), statut,
        float(row.get("target_aum_initial") or 0),
        float(row.get("revised_aum") or 0),
        float(row.get("funded_aum") or 0),
        raison_perte or None,
        str(row.get("concurrent_choisi") or "").strip() or None,
        nad_str,
        int(row["id"])
    ))
    conn.commit()
    conn.close()
    return True, None


# ---------------------------------------------------------------------------
# UPSERT — IMPORT CSV / EXCEL
# ---------------------------------------------------------------------------

def upsert_clients_from_df(df: pd.DataFrame) -> tuple:
    inserted, updated = 0, 0
    conn = get_connection()
    c = conn.cursor()
    df.columns = [col.lower().strip() for col in df.columns]

    for _, row in df.iterrows():
        nom = str(row.get("nom_client", "")).strip()
        if not nom:
            continue
        c.execute("SELECT id FROM clients WHERE nom_client=?", (nom,))
        existing = c.fetchone()
        if existing:
            c.execute(
                "UPDATE clients SET type_client=?, region=? WHERE nom_client=?",
                (str(row.get("type_client", "")), str(row.get("region", "")), nom)
            )
            updated += 1
        else:
            c.execute(
                "INSERT INTO clients (nom_client, type_client, region) VALUES (?,?,?)",
                (nom, str(row.get("type_client", "")), str(row.get("region", "")))
            )
            inserted += 1

    conn.commit()
    conn.close()
    return inserted, updated


def upsert_pipeline_from_df(df: pd.DataFrame) -> tuple:
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

        statut       = str(row.get("statut", "Prospect"))
        raison_perte = str(row.get("raison_perte", "") or "").strip() or None
        concurrent   = str(row.get("concurrent_choisi", "") or "").strip() or None
        next_action  = str(row.get("next_action_date", "") or "").strip() or None
        target_aum   = float(row.get("target_aum_initial", 0) or 0)
        revised_aum  = float(row.get("revised_aum", 0) or 0)
        funded_aum   = float(row.get("funded_aum", 0) or 0)

        c.execute("SELECT id FROM pipeline WHERE client_id=? AND fonds=?", (client_id, fonds))
        existing = c.fetchone()

        if existing:
            c.execute("""
                UPDATE pipeline SET statut=?, target_aum_initial=?, revised_aum=?,
                    funded_aum=?, raison_perte=?, concurrent_choisi=?,
                    next_action_date=?, updated_at=CURRENT_TIMESTAMP
                WHERE id=?
            """, (statut, target_aum, revised_aum, funded_aum,
                  raison_perte, concurrent, next_action, existing["id"]))
            updated += 1
        else:
            c.execute("""
                INSERT INTO pipeline
                    (client_id, fonds, statut, target_aum_initial, revised_aum,
                     funded_aum, raison_perte, concurrent_choisi, next_action_date)
                VALUES (?,?,?,?,?,?,?,?,?)
            """, (client_id, fonds, statut, target_aum, revised_aum,
                  funded_aum, raison_perte, concurrent, next_action))
            inserted += 1

    conn.commit()
    conn.close()
    return inserted, updated


# ---------------------------------------------------------------------------
# KPIs ANALYTIQUES
# ---------------------------------------------------------------------------

def get_kpis() -> dict:
    """Calcule les KPIs principaux pour le dashboard et le PDF."""
    conn = get_connection()
    c = conn.cursor()

    c.execute("SELECT COALESCE(SUM(funded_aum),0) FROM pipeline WHERE statut='Funded'")
    total_funded = float(c.fetchone()[0])

    c.execute("""
        SELECT COALESCE(SUM(revised_aum),0) FROM pipeline
        WHERE statut IN ('Prospect','Initial Pitch','Due Diligence','Soft Commit')
    """)
    pipeline_actif = float(c.fetchone()[0])

    c.execute("SELECT COUNT(*) FROM pipeline WHERE statut='Funded'")
    nb_funded = int(c.fetchone()[0])

    c.execute("SELECT COUNT(*) FROM pipeline WHERE statut='Lost'")
    nb_lost = int(c.fetchone()[0])

    taux_conv = (
        nb_funded / (nb_funded + nb_lost) * 100
        if (nb_funded + nb_lost) > 0 else 0.0
    )

    c.execute("""
        SELECT COUNT(*) FROM pipeline
        WHERE statut IN ('Prospect','Initial Pitch','Due Diligence','Soft Commit')
    """)
    nb_deals_actifs = int(c.fetchone()[0])

    c.execute("""
        SELECT c.type_client, COALESCE(SUM(p.funded_aum),0)
        FROM pipeline p JOIN clients c ON c.id=p.client_id
        WHERE p.statut='Funded'
        GROUP BY c.type_client
    """)
    aum_by_type = {r[0]: float(r[1]) for r in c.fetchall()}

    c.execute("""
        SELECT c.nom_client, c.type_client, c.region, p.fonds, p.funded_aum
        FROM pipeline p JOIN clients c ON c.id=p.client_id
        WHERE p.statut='Funded' AND p.funded_aum > 0
        ORDER BY p.funded_aum DESC LIMIT 10
    """)
    top_deals = [
        {k: (float(v) if k == "funded_aum" else str(v)) for k, v in dict(r).items()}
        for r in c.fetchall()
    ]

    c.execute("SELECT statut, COUNT(*) FROM pipeline GROUP BY statut")
    statut_repartition = {r[0]: int(r[1]) for r in c.fetchall()}

    c.execute("""
        SELECT fonds, COALESCE(SUM(funded_aum),0)
        FROM pipeline WHERE statut='Funded'
        GROUP BY fonds ORDER BY 2 DESC
    """)
    aum_by_fonds = {r[0]: float(r[1]) for r in c.fetchall()}

    conn.close()

    return {
        "total_funded":       total_funded,
        "pipeline_actif":     pipeline_actif,
        "taux_conversion":    round(taux_conv, 1),
        "nb_deals_actifs":    nb_deals_actifs,
        "nb_funded":          nb_funded,
        "nb_lost":            nb_lost,
        "aum_by_type":        aum_by_type,
        "top_deals":          top_deals,
        "statut_repartition": statut_repartition,
        "aum_by_fonds":       aum_by_fonds,
    }
