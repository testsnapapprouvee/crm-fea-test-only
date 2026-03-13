# =============================================================================
# database.py — Couche d'accès aux données SQLite
# CRM Asset Management - Charte Amundi
# =============================================================================

import sqlite3
import pandas as pd
from datetime import date, datetime, timedelta
import os

DB_PATH = os.path.join(os.path.dirname(__file__), "crm_asset_management.db")

# ---------------------------------------------------------------------------
# CONNEXION
# ---------------------------------------------------------------------------

def get_connection() -> sqlite3.Connection:
    """Retourne une connexion SQLite avec foreign keys activées."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


# ---------------------------------------------------------------------------
# INITIALISATION & DONNÉES FICTIVES
# ---------------------------------------------------------------------------

def init_db():
    """Crée les tables et insère les données fictives si la base est vide."""
    conn = get_connection()
    c = conn.cursor()

    # --- Table CLIENTS ---
    c.execute("""
        CREATE TABLE IF NOT EXISTS clients (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            nom_client    TEXT    NOT NULL UNIQUE,
            type_client   TEXT    NOT NULL
                          CHECK(type_client IN ('IFA','Wholesale','Instit','Family Office')),
            region        TEXT    NOT NULL
                          CHECK(region IN (
                              'GCC','EMEA','APAC','Nordics',
                              'Asia ex-Japan','North America','LatAm'
                          )),
            created_at    DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # --- Table PIPELINE ---
    c.execute("""
        CREATE TABLE IF NOT EXISTS pipeline (
            id                  INTEGER PRIMARY KEY AUTOINCREMENT,
            client_id           INTEGER NOT NULL,
            fonds               TEXT    NOT NULL
                                CHECK(fonds IN (
                                    'Global Value','International Fund','Income Builder',
                                    'Resilient Equity','Private Debt','Active ETFs'
                                )),
            statut              TEXT    NOT NULL DEFAULT 'Prospect'
                                CHECK(statut IN (
                                    'Prospect','Initial Pitch','Due Diligence',
                                    'Soft Commit','Funded','Paused','Lost','Redeemed'
                                )),
            target_aum_initial  REAL    DEFAULT 0,
            revised_aum         REAL    DEFAULT 0,
            funded_aum          REAL    DEFAULT 0,
            raison_perte        TEXT,
            concurrent_choisi   TEXT,
            next_action_date    DATE,
            created_at          DATETIME DEFAULT CURRENT_TIMESTAMP,
            updated_at          DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (client_id) REFERENCES clients(id) ON DELETE CASCADE
        )
    """)

    # --- Table ACTIVITES ---
    c.execute("""
        CREATE TABLE IF NOT EXISTS activites (
            id               INTEGER PRIMARY KEY AUTOINCREMENT,
            client_id        INTEGER NOT NULL,
            date             DATE    NOT NULL,
            notes            TEXT,
            type_interaction TEXT,
            created_at       DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (client_id) REFERENCES clients(id) ON DELETE CASCADE
        )
    """)

    conn.commit()

    # Insertion des données fictives uniquement si la base est vide
    c.execute("SELECT COUNT(*) FROM clients")
    if c.fetchone()[0] == 0:
        _insert_dummy_data(conn)

    conn.close()


def _insert_dummy_data(conn: sqlite3.Connection):
    """
    Insère 10 clients fictifs répartis sur différentes régions, types et fonds.
    Représente un pipeline réaliste pour une société d'Asset Management.
    """
    c = conn.cursor()
    today = date.today()

    # --- 10 Clients fictifs ---
    clients = [
        ("Al Rajhi Capital",        "IFA",          "GCC"),
        ("Emirates NBD AM",          "Wholesale",    "GCC"),
        ("Norges Bank Investment Mgmt","Instit",     "Nordics"),
        ("GIC Singapore",            "Instit",       "APAC"),
        ("Rothschild & Co Family Office","Family Office","EMEA"),
        ("BlackRock APAC Division",  "Wholesale",    "Asia ex-Japan"),
        ("JP Morgan Asset Management","Instit",      "North America"),
        ("ADIA – Abu Dhabi",         "Instit",       "GCC"),
        ("Lombard Odier FO Geneva",  "Family Office","EMEA"),
        ("BTG Pactual Wealth",       "Wholesale",    "LatAm"),
    ]
    c.executemany(
        "INSERT INTO clients (nom_client, type_client, region) VALUES (?,?,?)",
        clients
    )

    # --- Pipeline : un deal par client ---
    # (client_id, fonds, statut, target_init, revised, funded, raison, concurrent, next_action)
    pipeline = [
        (1, "Global Value",       "Funded",        50_000_000, 45_000_000, 42_000_000,
            None,        None,            (today + timedelta(days=15)).isoformat()),
        (2, "Income Builder",     "Soft Commit",   30_000_000, 28_000_000,  0,
            None,        None,            (today - timedelta(days=3)).isoformat()),   # overdue
        (3, "Private Debt",       "Due Diligence", 100_000_000,100_000_000, 0,
            None,        None,            (today + timedelta(days=7)).isoformat()),
        (4, "Resilient Equity",   "Funded",         75_000_000, 80_000_000,78_000_000,
            None,        None,            (today + timedelta(days=30)).isoformat()),
        (5, "International Fund", "Initial Pitch",  20_000_000, 20_000_000, 0,
            None,        None,            (today - timedelta(days=10)).isoformat()),  # overdue
        (6, "Active ETFs",        "Lost",           40_000_000, 35_000_000, 0,
            "Pricing",   "Vanguard",      (today + timedelta(days=60)).isoformat()),
        (7, "Global Value",       "Funded",        120_000_000,115_000_000,110_000_000,
            None,        None,            (today + timedelta(days=20)).isoformat()),
        (8, "Private Debt",       "Soft Commit",   200_000_000,180_000_000, 0,
            None,        None,            (today - timedelta(days=1)).isoformat()),   # overdue
        (9, "Income Builder",     "Paused",         15_000_000, 12_000_000, 0,
            "Macro",     "Internal",      (today + timedelta(days=90)).isoformat()),
        (10,"Resilient Equity",   "Due Diligence",  60_000_000, 60_000_000, 0,
            None,        None,            (today + timedelta(days=5)).isoformat()),
    ]
    c.executemany("""
        INSERT INTO pipeline
            (client_id, fonds, statut, target_aum_initial, revised_aum, funded_aum,
             raison_perte, concurrent_choisi, next_action_date)
        VALUES (?,?,?,?,?,?,?,?,?)
    """, pipeline)

    # --- Activités récentes ---
    activites = [
        (1, (today - timedelta(days=5)).isoformat(),  "Call de suivi post-investissement Q1",       "Call"),
        (2, (today - timedelta(days=2)).isoformat(),  "Envoi du term sheet révisé",                  "Email"),
        (3, (today - timedelta(days=1)).isoformat(),  "Réunion Due Diligence avec l'équipe risk",    "Meeting"),
        (4, (today - timedelta(days=8)).isoformat(),  "Confirmation d'investissement reçue",          "Email"),
        (7, (today - timedelta(days=3)).isoformat(),  "Review trimestrielle de la performance fonds", "Meeting"),
        (8, (today - timedelta(days=1)).isoformat(),  "Call de négociation des conditions LP",        "Call"),
        (5, (today - timedelta(days=6)).isoformat(),  "Présentation initiale du fonds International", "Meeting"),
        (10,(today - timedelta(days=4)).isoformat(),  "Envoi DDQ complété",                           "Email"),
    ]
    c.executemany("""
        INSERT INTO activites (client_id, date, notes, type_interaction)
        VALUES (?,?,?,?)
    """, activites)

    conn.commit()


# ---------------------------------------------------------------------------
# LECTURES — CLIENTS
# ---------------------------------------------------------------------------

def get_all_clients() -> pd.DataFrame:
    """Retourne tous les clients sous forme de DataFrame."""
    conn = get_connection()
    df = pd.read_sql_query(
        "SELECT id, nom_client, type_client, region FROM clients ORDER BY nom_client",
        conn
    )
    conn.close()
    return df


def get_client_options() -> dict:
    """Retourne un dict {nom_client: id} pour les selectbox."""
    df = get_all_clients()
    return dict(zip(df["nom_client"], df["id"]))


# ---------------------------------------------------------------------------
# LECTURES — PIPELINE
# ---------------------------------------------------------------------------

def get_pipeline_with_clients() -> pd.DataFrame:
    """
    Retourne le pipeline complet enrichi des infos client (JOIN).
    Colonnes: id, client_id, nom_client, type_client, region, fonds, statut,
              target_aum_initial, revised_aum, funded_aum, raison_perte,
              concurrent_choisi, next_action_date
    """
    conn = get_connection()
    df = pd.read_sql_query("""
        SELECT
            p.id,
            p.client_id,
            c.nom_client,
            c.type_client,
            c.region,
            p.fonds,
            p.statut,
            p.target_aum_initial,
            p.revised_aum,
            p.funded_aum,
            p.raison_perte,
            p.concurrent_choisi,
            p.next_action_date
        FROM pipeline p
        JOIN clients c ON c.id = p.client_id
        ORDER BY p.funded_aum DESC, p.revised_aum DESC
    """, conn)
    conn.close()
    # Conversion des colonnes AUM en float (sécurité)
    for col in ["target_aum_initial", "revised_aum", "funded_aum"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    return df


def get_overdue_actions() -> pd.DataFrame:
    """Retourne les deals avec une next_action_date dépassée (statut actif seulement)."""
    today_str = date.today().isoformat()
    conn = get_connection()
    df = pd.read_sql_query(f"""
        SELECT
            c.nom_client,
            c.type_client,
            p.fonds,
            p.statut,
            p.next_action_date
        FROM pipeline p
        JOIN clients c ON c.id = p.client_id
        WHERE p.next_action_date < '{today_str}'
          AND p.statut NOT IN ('Lost','Redeemed','Funded')
        ORDER BY p.next_action_date ASC
    """, conn)
    conn.close()
    return df


# ---------------------------------------------------------------------------
# LECTURES — ACTIVITÉS
# ---------------------------------------------------------------------------

def get_activities(client_id: int = None) -> pd.DataFrame:
    """Retourne les activités, optionnellement filtrées par client."""
    conn = get_connection()
    query = """
        SELECT a.id, c.nom_client, a.date, a.type_interaction, a.notes
        FROM activites a
        JOIN clients c ON c.id = a.client_id
    """
    if client_id:
        query += f" WHERE a.client_id = {int(client_id)}"
    query += " ORDER BY a.date DESC LIMIT 50"
    df = pd.read_sql_query(query, conn)
    conn.close()
    return df


# ---------------------------------------------------------------------------
# ÉCRITURE — CLIENTS
# ---------------------------------------------------------------------------

def add_client(nom_client: str, type_client: str, region: str) -> int:
    """Ajoute un client et retourne son ID."""
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
    """Met à jour un client existant."""
    conn = get_connection()
    conn.execute("""
        UPDATE clients SET nom_client=?, type_client=?, region=? WHERE id=?
    """, (nom_client.strip(), type_client, region, client_id))
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
    """Ajoute une entrée pipeline."""
    conn = get_connection()
    c = conn.cursor()
    c.execute("""
        INSERT INTO pipeline
            (client_id, fonds, statut, target_aum_initial, revised_aum, funded_aum,
             raison_perte, concurrent_choisi, next_action_date)
        VALUES (?,?,?,?,?,?,?,?,?)
    """, (client_id, fonds, statut, target_aum, revised_aum, funded_aum,
          raison_perte or None, concurrent_choisi or None, next_action_date))
    conn.commit()
    new_id = c.lastrowid
    conn.close()
    return new_id


def update_pipeline_row(row: dict):
    """
    Met à jour une ligne de pipeline depuis un dict.
    Valide la raison_perte si statut est Lost ou Paused.
    Retourne (True, None) si OK, (False, message_erreur) sinon.
    """
    if row.get("statut") in ("Lost", "Paused") and not row.get("raison_perte"):
        return False, f"⚠️ La raison de perte/pause est obligatoire pour le statut « {row['statut']} »."

    conn = get_connection()
    conn.execute("""
        UPDATE pipeline SET
            fonds=?, statut=?, target_aum_initial=?, revised_aum=?, funded_aum=?,
            raison_perte=?, concurrent_choisi=?, next_action_date=?,
            updated_at=CURRENT_TIMESTAMP
        WHERE id=?
    """, (
        row["fonds"], row["statut"],
        float(row.get("target_aum_initial") or 0),
        float(row.get("revised_aum") or 0),
        float(row.get("funded_aum") or 0),
        row.get("raison_perte") or None,
        row.get("concurrent_choisi") or None,
        row.get("next_action_date"),
        int(row["id"])
    ))
    conn.commit()
    conn.close()
    return True, None


def bulk_update_pipeline(df_edited: pd.DataFrame) -> list[str]:
    """
    Met à jour en masse le pipeline depuis un DataFrame édité.
    Retourne la liste des messages d'erreur.
    """
    errors = []
    for _, row in df_edited.iterrows():
        ok, msg = update_pipeline_row(row.to_dict())
        if not ok:
            errors.append(msg)
    return errors


# ---------------------------------------------------------------------------
# ÉCRITURE — ACTIVITÉS
# ---------------------------------------------------------------------------

def add_activity(client_id: int, date_str: str, notes: str, type_interaction: str):
    """Ajoute une activité."""
    conn = get_connection()
    conn.execute("""
        INSERT INTO activites (client_id, date, notes, type_interaction)
        VALUES (?,?,?,?)
    """, (client_id, date_str, notes, type_interaction))
    conn.commit()
    conn.close()


# ---------------------------------------------------------------------------
# UPSERT — IMPORT CSV/EXCEL
# ---------------------------------------------------------------------------

def upsert_clients_from_df(df: pd.DataFrame) -> tuple[int, int]:
    """
    Upsert clients depuis un DataFrame.
    Colonnes attendues : nom_client, type_client, region
    Retourne (nb_insérés, nb_mis_à_jour)
    """
    inserted, updated = 0, 0
    conn = get_connection()
    c = conn.cursor()

    # Normalisation des colonnes (case-insensitive)
    df.columns = [col.lower().strip() for col in df.columns]

    for _, row in df.iterrows():
        nom = str(row.get("nom_client", "")).strip()
        if not nom:
            continue
        c.execute("SELECT id FROM clients WHERE nom_client=?", (nom,))
        existing = c.fetchone()
        if existing:
            c.execute("""
                UPDATE clients SET type_client=?, region=? WHERE nom_client=?
            """, (row.get("type_client",""), row.get("region",""), nom))
            updated += 1
        else:
            c.execute("""
                INSERT INTO clients (nom_client, type_client, region) VALUES (?,?,?)
            """, (nom, row.get("type_client",""), row.get("region","")))
            inserted += 1

    conn.commit()
    conn.close()
    return inserted, updated


def upsert_pipeline_from_df(df: pd.DataFrame) -> tuple[int, int]:
    """
    Upsert pipeline depuis un DataFrame.
    Colonnes attendues : nom_client (ou client_id), fonds, statut, ... 
    Retourne (nb_insérés, nb_mis_à_jour)
    """
    inserted, updated = 0, 0
    conn = get_connection()
    c = conn.cursor()
    df.columns = [col.lower().strip() for col in df.columns]

    for _, row in df.iterrows():
        # Résolution du client_id depuis nom_client
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

        # Vérification : existe-t-il déjà un deal (client_id + fonds) ?
        c.execute("""
            SELECT id FROM pipeline WHERE client_id=? AND fonds=?
        """, (client_id, fonds))
        existing = c.fetchone()

        statut         = str(row.get("statut", "Prospect"))
        raison_perte   = str(row.get("raison_perte", "")) or None
        concurrent     = str(row.get("concurrent_choisi", "")) or None
        next_action    = str(row.get("next_action_date", "")) or None
        target_aum     = float(row.get("target_aum_initial", 0) or 0)
        revised_aum    = float(row.get("revised_aum", 0) or 0)
        funded_aum     = float(row.get("funded_aum", 0) or 0)

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
    """
    Calcule les KPIs principaux pour le dashboard.
    Retourne un dict avec les métriques clés.
    """
    conn = get_connection()
    c = conn.cursor()

    # AUM total financé
    c.execute("SELECT COALESCE(SUM(funded_aum),0) FROM pipeline WHERE statut='Funded'")
    total_funded = c.fetchone()[0]

    # Pipeline actif (Soft Commit + Due Diligence + Initial Pitch)
    c.execute("""
        SELECT COALESCE(SUM(revised_aum),0) FROM pipeline
        WHERE statut IN ('Prospect','Initial Pitch','Due Diligence','Soft Commit')
    """)
    pipeline_actif = c.fetchone()[0]

    # Taux de conversion (Funded / (Funded + Lost))
    c.execute("SELECT COUNT(*) FROM pipeline WHERE statut='Funded'")
    nb_funded = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM pipeline WHERE statut='Lost'")
    nb_lost = c.fetchone()[0]
    taux_conv = (nb_funded / (nb_funded + nb_lost) * 100) if (nb_funded + nb_lost) > 0 else 0

    # Nb deals actifs
    c.execute("""
        SELECT COUNT(*) FROM pipeline
        WHERE statut IN ('Prospect','Initial Pitch','Due Diligence','Soft Commit')
    """)
    nb_deals_actifs = c.fetchone()[0]

    # AUM par type de client (Funded)
    c.execute("""
        SELECT c.type_client, COALESCE(SUM(p.funded_aum),0) as aum
        FROM pipeline p JOIN clients c ON c.id=p.client_id
        WHERE p.statut='Funded'
        GROUP BY c.type_client
    """)
    aum_by_type = {r[0]: r[1] for r in c.fetchall()}

    # Top deals Funded
    c.execute("""
        SELECT c.nom_client, c.type_client, c.region, p.fonds, p.funded_aum
        FROM pipeline p JOIN clients c ON c.id=p.client_id
        WHERE p.statut='Funded' AND p.funded_aum > 0
        ORDER BY p.funded_aum DESC
        LIMIT 10
    """)
    top_deals = c.fetchall()

    # Répartition par statut
    c.execute("SELECT statut, COUNT(*) as nb FROM pipeline GROUP BY statut")
    statut_repartition = {r[0]: r[1] for r in c.fetchall()}

    # AUM par fonds
    c.execute("""
        SELECT fonds, COALESCE(SUM(funded_aum),0) as aum
        FROM pipeline WHERE statut='Funded'
        GROUP BY fonds ORDER BY aum DESC
    """)
    aum_by_fonds = {r[0]: r[1] for r in c.fetchall()}

    conn.close()

    return {
        "total_funded":       total_funded,
        "pipeline_actif":     pipeline_actif,
        "taux_conversion":    round(taux_conv, 1),
        "nb_deals_actifs":    nb_deals_actifs,
        "nb_funded":          nb_funded,
        "nb_lost":            nb_lost,
        "aum_by_type":        aum_by_type,
        "top_deals":          [dict(r) for r in top_deals],
        "statut_repartition": statut_repartition,
        "aum_by_fonds":       aum_by_fonds,
    }
