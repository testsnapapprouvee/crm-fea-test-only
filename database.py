# =============================================================================
# database.py — Couche d'acces aux donnees SQLite
# CRM Asset Management — Edition Enterprise
# Staff Engineer refactoring : typage strict, audit trail, sales tracking
# =============================================================================

import sqlite3
import pandas as pd
import numpy as np
from datetime import date, datetime, timedelta
from typing import Optional
import os

DB_PATH = os.path.join(os.path.dirname(__file__), "crm_v3.db")

_AUM_COLS           = ["target_aum_initial", "revised_aum", "funded_aum"]
_TEXT_NULLABLE_COLS = ["raison_perte", "concurrent_choisi"]
_AUDIT_CHAMPS       = ["statut", "fonds", "target_aum_initial",
                       "revised_aum", "funded_aum", "raison_perte"]


# ---------------------------------------------------------------------------
# CONNEXION
# ---------------------------------------------------------------------------

def get_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


# ---------------------------------------------------------------------------
# NETTOYAGE DE TYPES — CONTRAT DE DONNEES (anti-crash StreamlitAPIException)
# ---------------------------------------------------------------------------

def _clean_pipeline_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normalise tous les types du DataFrame pipeline avant exposition a Streamlit.
    Garanties : AUM=float64, next_action_date=datetime.date|None,
    textes nullable=str, id/client_id=int64.
    """
    if df.empty:
        return df.copy()

    df = df.copy()

    for col in _AUM_COLS:
        if col in df.columns:
            df[col] = (
                pd.to_numeric(df[col], errors="coerce")
                .fillna(0.0)
                .astype("float64")
            )

    if "next_action_date" in df.columns:
        parsed = pd.to_datetime(df["next_action_date"], errors="coerce")
        df["next_action_date"] = [
            ts.date() if pd.notna(ts) else None for ts in parsed
        ]

    for col in _TEXT_NULLABLE_COLS:
        if col in df.columns:
            df[col] = df[col].apply(
                lambda v: ""
                if (v is None or (isinstance(v, float) and np.isnan(v)))
                else str(v)
            )

    for col in ["id", "client_id"]:
        if col in df.columns:
            df[col] = (
                pd.to_numeric(df[col], errors="coerce")
                .fillna(0).astype("int64")
            )

    return df


# ---------------------------------------------------------------------------
# INITIALISATION DB — MIGRATION ADDITIVE SANS PERTE DE DONNEES
# ---------------------------------------------------------------------------

def init_db():
    """
    Cree les tables et applique les migrations additives (ALTER TABLE).
    Idempotent : peut etre appele plusieurs fois sans erreur.
    """
    conn = get_connection()
    c = conn.cursor()

    # --- Table CLIENTS ---
    c.execute("""
        CREATE TABLE IF NOT EXISTS clients (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            nom_client  TEXT NOT NULL UNIQUE,
            type_client TEXT NOT NULL
                        CHECK(type_client IN
                              ('IFA','Wholesale','Instit','Family Office')),
            region      TEXT NOT NULL
                        CHECK(region IN ('GCC','EMEA','APAC','Nordics',
                                         'Asia ex-Japan','North America','LatAm')),
            created_at  DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # --- Table PIPELINE ---
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
                                                'Soft Commit','Funded','Paused',
                                                'Lost','Redeemed')),
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

    # Migration additive : ajouter sales_owner si la table existait sans elle
    try:
        c.execute(
            "ALTER TABLE pipeline ADD COLUMN "
            "sales_owner TEXT NOT NULL DEFAULT 'Non assigne'"
        )
        conn.commit()
    except sqlite3.OperationalError:
        pass  # Colonne deja presente

    # --- Table ACTIVITES ---
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

    # --- Table AUDIT_LOG ---
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
        _insert_dummy_data(conn)

    conn.close()


def _insert_dummy_data(conn: sqlite3.Connection):
    """10 clients fictifs avec 3 sales owners pour les metriques Sales Tracking."""
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

    # (client_id, fonds, statut, target, revised, funded,
    #  raison, concurrent, next_action, sales_owner)
    pipeline = [
        (1,  "Global Value",       "Funded",        50_000_000,  45_000_000, 42_000_000,
         None, None,
         (today + timedelta(days=15)).isoformat(), "Sophie Laurent"),
        (2,  "Income Builder",     "Soft Commit",   30_000_000,  28_000_000,          0,
         None, None,
         (today - timedelta(days=3)).isoformat(),  "Marc Dupont"),
        (3,  "Private Debt",       "Due Diligence", 100_000_000, 100_000_000,         0,
         None, None,
         (today + timedelta(days=7)).isoformat(),  "Sophie Laurent"),
        (4,  "Resilient Equity",   "Funded",         75_000_000,  80_000_000, 78_000_000,
         None, None,
         (today + timedelta(days=30)).isoformat(), "Karim Belhadj"),
        (5,  "International Fund", "Initial Pitch",  20_000_000,  20_000_000,          0,
         None, None,
         (today - timedelta(days=10)).isoformat(), "Marc Dupont"),
        (6,  "Active ETFs",        "Lost",           40_000_000,  35_000_000,          0,
         "Pricing", "Vanguard",
         (today + timedelta(days=60)).isoformat(), "Karim Belhadj"),
        (7,  "Global Value",       "Funded",        120_000_000, 115_000_000, 110_000_000,
         None, None,
         (today + timedelta(days=20)).isoformat(), "Sophie Laurent"),
        (8,  "Private Debt",       "Soft Commit",   200_000_000, 180_000_000,          0,
         None, None,
         (today - timedelta(days=1)).isoformat(),  "Karim Belhadj"),
        (9,  "Income Builder",     "Paused",         15_000_000,  12_000_000,          0,
         "Macro", "Internal",
         (today + timedelta(days=90)).isoformat(), "Marc Dupont"),
        (10, "Resilient Equity",   "Due Diligence",  60_000_000,  60_000_000,          0,
         None, None,
         (today + timedelta(days=5)).isoformat(),  "Sophie Laurent"),
    ]
    c.executemany("""
        INSERT INTO pipeline
            (client_id, fonds, statut, target_aum_initial, revised_aum, funded_aum,
             raison_perte, concurrent_choisi, next_action_date, sales_owner)
        VALUES (?,?,?,?,?,?,?,?,?,?)
    """, pipeline)

    activites = [
        (1,  (today - timedelta(days=5)).isoformat(),
         "Suivi post-investissement Q1", "Call"),
        (2,  (today - timedelta(days=2)).isoformat(),
         "Envoi du term sheet revise", "Email"),
        (3,  (today - timedelta(days=1)).isoformat(),
         "Reunion Due Diligence — equipe risk", "Meeting"),
        (4,  (today - timedelta(days=8)).isoformat(),
         "Confirmation investissement recue", "Email"),
        (7,  (today - timedelta(days=3)).isoformat(),
         "Review trimestrielle performance", "Meeting"),
        (8,  (today - timedelta(days=1)).isoformat(),
         "Negociation conditions LP", "Call"),
        (5,  (today - timedelta(days=6)).isoformat(),
         "Presentation initiale International Fund", "Meeting"),
        (10, (today - timedelta(days=4)).isoformat(),
         "Envoi DDQ complete", "Email"),
    ]
    c.executemany(
        "INSERT INTO activites (client_id, date, notes, type_interaction) "
        "VALUES (?,?,?,?)",
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
    """Dict {nom_client: id Python int} pour les selectbox."""
    df = get_all_clients()
    return {str(n): int(i) for n, i in zip(df["nom_client"], df["id"])}


# ---------------------------------------------------------------------------
# LECTURES — PIPELINE
# ---------------------------------------------------------------------------

def get_pipeline_with_clients() -> pd.DataFrame:
    """Pipeline complet avec JOIN clients. Tous les types sont normalises."""
    conn = get_connection()
    df = pd.read_sql_query("""
        SELECT
            p.id, p.client_id,
            c.nom_client, c.type_client, c.region,
            p.fonds, p.statut,
            p.target_aum_initial, p.revised_aum, p.funded_aum,
            p.raison_perte, p.concurrent_choisi,
            p.next_action_date, p.sales_owner
        FROM pipeline p
        JOIN clients c ON c.id = p.client_id
        ORDER BY p.funded_aum DESC, p.revised_aum DESC
    """, conn)
    conn.close()
    return _clean_pipeline_df(df)


def get_pipeline_row_by_id(pipeline_id: int) -> Optional[dict]:
    """
    Retourne une ligne pipeline sous forme de dict Python natif strict.
    Utilise par le formulaire Master-Detail pour pre-remplir les champs.
    """
    conn = get_connection()
    df = pd.read_sql_query("""
        SELECT
            p.id, p.client_id, c.nom_client, c.type_client, c.region,
            p.fonds, p.statut, p.target_aum_initial, p.revised_aum, p.funded_aum,
            p.raison_perte, p.concurrent_choisi, p.next_action_date, p.sales_owner
        FROM pipeline p
        JOIN clients c ON c.id = p.client_id
        WHERE p.id = ?
    """, conn, params=(int(pipeline_id),))
    conn.close()

    if df.empty:
        return None

    df = _clean_pipeline_df(df)
    row = df.iloc[0].to_dict()

    for col in _AUM_COLS:
        row[col] = float(row.get(col) or 0.0)
    row["id"]           = int(row["id"])
    row["client_id"]    = int(row["client_id"])
    row["sales_owner"]  = str(row.get("sales_owner") or "Non assigne")

    for col in _TEXT_NULLABLE_COLS:
        row[col] = str(row.get(col) or "")

    nad = row.get("next_action_date")
    if not isinstance(nad, date):
        row["next_action_date"] = date.today() + timedelta(days=14)

    return row


def get_overdue_actions() -> pd.DataFrame:
    today_str = date.today().isoformat()
    conn = get_connection()
    df = pd.read_sql_query(f"""
        SELECT
            c.nom_client, c.type_client, p.fonds, p.statut,
            p.next_action_date, p.sales_owner
        FROM pipeline p
        JOIN clients c ON c.id = p.client_id
        WHERE p.next_action_date < '{today_str}'
          AND p.statut NOT IN ('Lost','Redeemed','Funded')
        ORDER BY p.next_action_date ASC
    """, conn)
    conn.close()
    return _clean_pipeline_df(df)


# ---------------------------------------------------------------------------
# LECTURES — AUDIT LOG
# ---------------------------------------------------------------------------

def get_audit_log(pipeline_id: int) -> pd.DataFrame:
    """Retourne l'historique complet des modifications d'un deal."""
    conn = get_connection()
    df = pd.read_sql_query("""
        SELECT
            champ_modifie    AS "Champ",
            ancienne_valeur  AS "Ancienne valeur",
            nouvelle_valeur  AS "Nouvelle valeur",
            modified_by      AS "Modifie par",
            date_modification AS "Date"
        FROM audit_log
        WHERE pipeline_id = ?
        ORDER BY date_modification DESC
    """, conn, params=(int(pipeline_id),))
    conn.close()
    return df


# ---------------------------------------------------------------------------
# LECTURES — SALES TRACKING
# ---------------------------------------------------------------------------

def get_sales_metrics() -> pd.DataFrame:
    """
    Metriques aggregees par sales_owner :
    nb_deals_total, nb_funded, aum_funded, pipeline_value, nb_overdue
    """
    today_str = date.today().isoformat()
    conn = get_connection()
    df = pd.read_sql_query(f"""
        SELECT
            p.sales_owner                                            AS "Commercial",
            COUNT(*)                                                  AS "Nb Deals",
            COUNT(CASE WHEN p.statut='Funded' THEN 1 END)            AS "Funded",
            COUNT(CASE WHEN p.statut IN
                  ('Prospect','Initial Pitch','Due Diligence','Soft Commit')
                  THEN 1 END)                                         AS "Actifs",
            COUNT(CASE WHEN p.statut='Lost' THEN 1 END)              AS "Perdus",
            COALESCE(SUM(
                CASE WHEN p.statut='Funded' THEN p.funded_aum END
            ), 0)                                                     AS "AUM Finance",
            COALESCE(SUM(
                CASE WHEN p.statut IN
                ('Prospect','Initial Pitch','Due Diligence','Soft Commit')
                THEN p.revised_aum END
            ), 0)                                                     AS "Pipeline Actif",
            COUNT(CASE WHEN p.next_action_date < '{today_str}'
                  AND p.statut NOT IN ('Lost','Redeemed','Funded')
                  THEN 1 END)                                         AS "Actions en retard"
        FROM pipeline p
        GROUP BY p.sales_owner
        ORDER BY "AUM Finance" DESC
    """, conn)
    conn.close()
    for col in ["AUM Finance", "Pipeline Actif"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
    return df


def get_next_actions_by_sales(days_ahead: int = 30) -> pd.DataFrame:
    """
    Prochaines actions des 'days_ahead' prochains jours (+ overdue),
    groupees par sales_owner, statuts actifs uniquement.
    """
    today_str    = date.today().isoformat()
    horizon_str  = (date.today() + timedelta(days=days_ahead)).isoformat()
    conn = get_connection()
    df = pd.read_sql_query(f"""
        SELECT
            p.sales_owner,
            c.nom_client,
            p.fonds,
            p.statut,
            p.next_action_date,
            p.revised_aum
        FROM pipeline p
        JOIN clients c ON c.id = p.client_id
        WHERE p.statut NOT IN ('Lost','Redeemed','Funded')
          AND (
                p.next_action_date <= '{horizon_str}'
                OR p.next_action_date < '{today_str}'
          )
        ORDER BY p.sales_owner, p.next_action_date ASC
    """, conn)
    conn.close()
    return _clean_pipeline_df(df)


def get_sales_owners() -> list:
    """Liste triee de tous les sales owners distincts."""
    conn = get_connection()
    c = conn.cursor()
    c.execute(
        "SELECT DISTINCT sales_owner FROM pipeline "
        "WHERE sales_owner IS NOT NULL "
        "ORDER BY sales_owner"
    )
    owners = [r[0] for r in c.fetchall()]
    conn.close()
    return owners


# ---------------------------------------------------------------------------
# LECTURES — ACTIVITES
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
# ECRITURE — CLIENTS
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


def update_client(client_id: int, nom_client: str,
                  type_client: str, region: str):
    conn = get_connection()
    conn.execute(
        "UPDATE clients SET nom_client=?, type_client=?, region=? WHERE id=?",
        (nom_client.strip(), type_client, region, int(client_id))
    )
    conn.commit()
    conn.close()


# ---------------------------------------------------------------------------
# ECRITURE — PIPELINE (avec AUDIT TRAIL)
# ---------------------------------------------------------------------------

def add_pipeline_entry(
    client_id: int, fonds: str, statut: str,
    target_aum: float, revised_aum: float, funded_aum: float,
    raison_perte: str, concurrent_choisi: str,
    next_action_date: str, sales_owner: str = "Non assigne"
) -> int:
    conn = get_connection()
    c = conn.cursor()
    c.execute("""
        INSERT INTO pipeline
            (client_id, fonds, statut, target_aum_initial, revised_aum, funded_aum,
             raison_perte, concurrent_choisi, next_action_date, sales_owner)
        VALUES (?,?,?,?,?,?,?,?,?,?)
    """, (
        int(client_id), fonds, statut,
        float(target_aum or 0), float(revised_aum or 0), float(funded_aum or 0),
        raison_perte.strip() or None if raison_perte else None,
        concurrent_choisi.strip() or None if concurrent_choisi else None,
        next_action_date,
        sales_owner.strip() or "Non assigne",
    ))
    conn.commit()
    new_id = c.lastrowid
    conn.close()
    return new_id


def update_pipeline_row(row: dict, modified_by: str = "utilisateur") -> tuple:
    """
    Met a jour une ligne pipeline et alimente l'audit log.
    Validation : raison_perte obligatoire si statut Lost/Paused.
    Retourne (True, None) si OK, (False, message_erreur) sinon.
    """
    statut       = str(row.get("statut", ""))
    raison_perte = str(row.get("raison_perte") or "").strip()

    if statut in ("Lost", "Paused") and not raison_perte:
        return False, (
            f"La raison de perte/pause est obligatoire "
            f"pour le statut « {statut} »."
        )

    # Normalisation next_action_date -> str ISO
    nad = row.get("next_action_date")
    if isinstance(nad, (date, datetime)):
        nad_str = nad.strftime("%Y-%m-%d")
    elif isinstance(nad, str) and nad.strip():
        nad_str = nad.strip()
    else:
        nad_str = None

    pipeline_id  = int(row["id"])
    new_fonds    = str(row["fonds"])
    new_target   = float(row.get("target_aum_initial") or 0)
    new_revised  = float(row.get("revised_aum") or 0)
    new_funded   = float(row.get("funded_aum") or 0)
    new_raison   = raison_perte or None
    new_concurrent = str(row.get("concurrent_choisi") or "").strip() or None
    new_sales    = str(row.get("sales_owner") or "Non assigne").strip()

    conn = get_connection()
    c = conn.cursor()

    # --- Lecture de l'ancienne ligne pour l'audit trail ---
    c.execute("""
        SELECT fonds, statut, target_aum_initial, revised_aum,
               funded_aum, raison_perte
        FROM pipeline WHERE id = ?
    """, (pipeline_id,))
    old_row = c.fetchone()

    # --- INSERT dans audit_log pour chaque champ modifie ---
    if old_row:
        old_vals = {
            "fonds":             str(old_row["fonds"] or ""),
            "statut":            str(old_row["statut"] or ""),
            "target_aum_initial": float(old_row["target_aum_initial"] or 0),
            "revised_aum":       float(old_row["revised_aum"] or 0),
            "funded_aum":        float(old_row["funded_aum"] or 0),
            "raison_perte":      str(old_row["raison_perte"] or ""),
        }
        new_vals = {
            "fonds":             new_fonds,
            "statut":            statut,
            "target_aum_initial": new_target,
            "revised_aum":       new_revised,
            "funded_aum":        new_funded,
            "raison_perte":      str(new_raison or ""),
        }

        for champ in _AUDIT_CHAMPS:
            old_v = old_vals.get(champ)
            new_v = new_vals.get(champ)
            # Comparaison numerique avec tolerance pour les AUM
            if champ in ("target_aum_initial", "revised_aum", "funded_aum"):
                changed = abs(float(old_v or 0) - float(new_v or 0)) > 0.01
            else:
                changed = str(old_v) != str(new_v)

            if changed:
                c.execute("""
                    INSERT INTO audit_log
                        (pipeline_id, champ_modifie, ancienne_valeur,
                         nouvelle_valeur, modified_by)
                    VALUES (?,?,?,?,?)
                """, (
                    pipeline_id, champ,
                    str(old_v), str(new_v),
                    modified_by
                ))

    # --- UPDATE ---
    c.execute("""
        UPDATE pipeline SET
            fonds=?, statut=?,
            target_aum_initial=?, revised_aum=?, funded_aum=?,
            raison_perte=?, concurrent_choisi=?,
            next_action_date=?, sales_owner=?,
            updated_at=CURRENT_TIMESTAMP
        WHERE id=?
    """, (
        new_fonds, statut,
        new_target, new_revised, new_funded,
        new_raison, new_concurrent,
        nad_str, new_sales,
        pipeline_id
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
                (str(row.get("type_client", "")),
                 str(row.get("region", "")), nom)
            )
            updated += 1
        else:
            c.execute(
                "INSERT INTO clients (nom_client, type_client, region) "
                "VALUES (?,?,?)",
                (nom, str(row.get("type_client", "")),
                 str(row.get("region", "")))
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

        statut      = str(row.get("statut", "Prospect"))
        raison      = str(row.get("raison_perte", "") or "").strip() or None
        concurrent  = str(row.get("concurrent_choisi", "") or "").strip() or None
        next_action = str(row.get("next_action_date", "") or "").strip() or None
        target_aum  = float(row.get("target_aum_initial", 0) or 0)
        revised_aum = float(row.get("revised_aum", 0) or 0)
        funded_aum  = float(row.get("funded_aum", 0) or 0)
        sales_owner = str(row.get("sales_owner", "Non assigne") or "Non assigne")

        c.execute(
            "SELECT id FROM pipeline WHERE client_id=? AND fonds=?",
            (client_id, fonds)
        )
        existing = c.fetchone()

        if existing:
            c.execute("""
                UPDATE pipeline SET statut=?, target_aum_initial=?,
                    revised_aum=?, funded_aum=?, raison_perte=?,
                    concurrent_choisi=?, next_action_date=?, sales_owner=?,
                    updated_at=CURRENT_TIMESTAMP
                WHERE id=?
            """, (statut, target_aum, revised_aum, funded_aum,
                  raison, concurrent, next_action, sales_owner,
                  existing["id"]))
            updated += 1
        else:
            c.execute("""
                INSERT INTO pipeline
                    (client_id, fonds, statut, target_aum_initial, revised_aum,
                     funded_aum, raison_perte, concurrent_choisi,
                     next_action_date, sales_owner)
                VALUES (?,?,?,?,?,?,?,?,?,?)
            """, (client_id, fonds, statut, target_aum, revised_aum,
                  funded_aum, raison, concurrent, next_action, sales_owner))
            inserted += 1

    conn.commit()
    conn.close()
    return inserted, updated


# ---------------------------------------------------------------------------
# KPIs ANALYTIQUES
# ---------------------------------------------------------------------------

def get_kpis() -> dict:
    conn = get_connection()
    c = conn.cursor()

    c.execute(
        "SELECT COALESCE(SUM(funded_aum),0) FROM pipeline WHERE statut='Funded'"
    )
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
        SELECT c.nom_client, c.type_client, c.region, p.fonds,
               p.funded_aum, p.sales_owner
        FROM pipeline p JOIN clients c ON c.id=p.client_id
        WHERE p.statut='Funded' AND p.funded_aum > 0
        ORDER BY p.funded_aum DESC LIMIT 10
    """)
    top_deals = [
        {k: (float(v) if k == "funded_aum" else str(v))
         for k, v in dict(r).items()}
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


# ---------------------------------------------------------------------------
# AJOUT ACTIVITE
# ---------------------------------------------------------------------------

def add_activity(client_id: int, date_str: str,
                 notes: str, type_interaction: str):
    conn = get_connection()
    conn.execute(
        "INSERT INTO activites (client_id, date, notes, type_interaction) "
        "VALUES (?,?,?,?)",
        (int(client_id), date_str, notes, type_interaction)
    )
    conn.commit()
    conn.close()
