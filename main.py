# =============================================================================
# main.py — Meridian CRM · FastAPI backend
# Serves frontend/revolut_crm_dashboard_v2.html and exposes /api/* JSON endpoints
# backed by the existing SQLite database (database.py)
# =============================================================================

import os
from datetime import date, datetime, timedelta
from typing import Optional

import numpy as np
import pandas as pd
from fastapi import FastAPI, HTTPException
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

import database as db

ROOT_DIR     = os.path.dirname(os.path.abspath(__file__))
FRONTEND_DIR = os.path.join(ROOT_DIR, "frontend")
INDEX_PATH   = os.path.join(FRONTEND_DIR, "revolut_crm_dashboard_v2.html")

app = FastAPI(
    title="Meridian CRM API",
    description="JSON REST API serving the Meridian premium dashboard.",
    version="1.0.0",
)

# Initialise the SQLite schema on startup
db.init_db()

# Static assets directory (CSS images later, etc.)
if os.path.isdir(FRONTEND_DIR):
    app.mount("/static", StaticFiles(directory=FRONTEND_DIR), name="static")


# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------

def _smart_aum(row) -> float:
    """Pipeline AUM: revised if > 0, else target. Funded uses funded_aum."""
    statut = str(row.get("statut", ""))
    if statut == "Funded":
        return float(row.get("funded_aum", 0) or 0)
    rev = float(row.get("revised_aum", 0) or 0)
    if rev > 0:
        return rev
    return float(row.get("target_aum_initial", 0) or 0)


def _safe_dict(d):
    """Convert numpy/pandas types to JSON-friendly natives."""
    out = {}
    for k, v in d.items():
        if isinstance(v, (np.integer,)):
            out[k] = int(v)
        elif isinstance(v, (np.floating,)):
            out[k] = float(v)
        elif isinstance(v, (np.bool_,)):
            out[k] = bool(v)
        elif isinstance(v, (pd.Timestamp,)):
            out[k] = v.isoformat()
        elif isinstance(v, (date, datetime)):
            out[k] = v.isoformat()
        else:
            out[k] = v
    return out


# ---------------------------------------------------------------------------
# ROOT — serve the HTML
# ---------------------------------------------------------------------------

@app.get("/", response_class=HTMLResponse)
def root():
    if not os.path.exists(INDEX_PATH):
        raise HTTPException(status_code=404, detail="Frontend index not found.")
    with open(INDEX_PATH, "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())


@app.get("/health")
def health():
    return {"status": "ok", "service": "meridian-crm-api"}


# ---------------------------------------------------------------------------
# /api/kpis — header KPIs
# ---------------------------------------------------------------------------

@app.get("/api/kpis")
def api_kpis():
    k = db.get_kpis()
    return {
        "total_funded":      float(k.get("total_funded", 0)),
        "pipeline_actif":    float(k.get("pipeline_actif", 0)),
        "weighted_pipeline": float(k.get("weighted_pipeline", 0)),
        "taux_conversion":   float(k.get("taux_conversion", 0)),
        "nb_funded":         int(k.get("nb_funded", 0)),
        "nb_lost":           int(k.get("nb_lost", 0)),
        "nb_paused":         int(k.get("nb_paused", 0)),
        "nb_deals_actifs":   int(k.get("nb_deals_actifs", 0)),
        "statut_repartition": k.get("statut_repartition", {}),
    }


# ---------------------------------------------------------------------------
# /api/deals — list of pipeline deals (joined with clients)
# ---------------------------------------------------------------------------

@app.get("/api/deals")
def api_deals():
    df = db.get_pipeline_with_clients()
    if df.empty:
        return []
    items = []
    for _, r in df.iterrows():
        statut = str(r.get("statut", ""))
        aum = float(r.get("funded_aum", 0) or 0) if statut == "Funded" else _smart_aum(r)
        items.append(_safe_dict({
            "id":          int(r["id"]),
            "client_id":   int(r["client_id"]),
            "client":      str(r["nom_client"]),
            "type_client": str(r.get("type_client", "")),
            "region":      str(r.get("region", "")),
            "country":     str(r.get("country", "")),
            "fonds":       str(r.get("fonds", "")),
            "statut":      statut,
            "aum":         aum,
            "funded_aum":  float(r.get("funded_aum", 0) or 0),
            "revised_aum": float(r.get("revised_aum", 0) or 0),
            "target_aum":  float(r.get("target_aum_initial", 0) or 0),
            "sales_owner": str(r.get("sales_owner", "")),
            "next_action_date": r.get("next_action_date").isoformat()
                if isinstance(r.get("next_action_date"), date) else None,
            "closing_probability": float(r.get("closing_probability", 50) or 50),
        }))
    items.sort(key=lambda x: x["aum"], reverse=True)
    return items


# ---------------------------------------------------------------------------
# /api/clients — list of all clients with aggregate AUM
# ---------------------------------------------------------------------------

@app.get("/api/clients")
def api_clients():
    df_clients = db.get_client_hierarchy()
    df_pipe    = db.get_pipeline_with_clients()
    aum_by_client = {}
    if not df_pipe.empty:
        for cid, grp in df_pipe.groupby("client_id"):
            funded = float(grp[grp["statut"] == "Funded"]["funded_aum"].sum())
            active = grp[grp["statut"].isin(
                ["Prospect","Initial Pitch","Due Diligence","Soft Commit"])]
            active_aum = float(active.apply(_smart_aum, axis=1).sum()) if not active.empty else 0.0
            aum_by_client[int(cid)] = funded + active_aum
    out = []
    for _, r in df_clients.iterrows():
        cid = int(r["id"])
        out.append(_safe_dict({
            "id":          cid,
            "nom_client":  str(r["nom_client"]),
            "type_client": str(r.get("type_client", "")),
            "region":      str(r.get("region", "")),
            "country":     str(r.get("country", "")),
            "tier":        str(r.get("tier", "")),
            "kyc_status":  str(r.get("kyc_status", "")),
            "aum":         aum_by_client.get(cid, 0.0),
        }))
    out.sort(key=lambda c: c["aum"], reverse=True)
    return out


# ---------------------------------------------------------------------------
# /api/clients/{id} — 360° detail
# ---------------------------------------------------------------------------

@app.get("/api/clients/{client_id}")
def api_client_detail(client_id: int):
    df_clients = db.get_client_hierarchy()
    row = df_clients[df_clients["id"] == client_id]
    if row.empty:
        raise HTTPException(status_code=404, detail="Client not found")
    c = row.iloc[0].to_dict()

    contacts_df = db.get_contacts(client_id)
    activities_df = db.get_activities(client_id=client_id)
    df_pipe = db.get_pipeline_with_clients()
    deals = df_pipe[df_pipe["client_id"] == client_id]

    funded_aum = float(deals[deals["statut"] == "Funded"]["funded_aum"].sum()) \
        if not deals.empty else 0.0
    active = deals[deals["statut"].isin(
        ["Prospect","Initial Pitch","Due Diligence","Soft Commit"])] \
        if not deals.empty else deals
    active_aum = float(active.apply(_smart_aum, axis=1).sum()) if not active.empty else 0.0

    fonds_invested = sorted(deals[deals["statut"] == "Funded"]["fonds"].unique().tolist()) \
        if not deals.empty else []

    dominant_statut = "—"
    if not deals.empty:
        cnt = deals["statut"].value_counts()
        dominant_statut = str(cnt.index[0])

    primary_contact = None
    if not contacts_df.empty:
        primary = contacts_df[contacts_df["is_primary"] == 1]
        if not primary.empty:
            pc = primary.iloc[0]
            primary_contact = "{} {}".format(pc.get("prenom", ""), pc.get("nom", "")).strip()
        elif not contacts_df.empty:
            pc = contacts_df.iloc[0]
            primary_contact = "{} {}".format(pc.get("prenom", ""), pc.get("nom", "")).strip()

    network = db.get_client_network(client_id)
    fund_links_clean = []
    for fl in network.get("fund_links", []):
        fund_links_clean.append(_safe_dict({
            "fonds":     str(fl.get("fonds", "")),
            "client_id": int(fl.get("client_id", 0)),
            "statut":    str(fl.get("statut", "")),
        }))
    subs_clean = []
    for s in network.get("subsidiaries", []):
        subs_clean.append(_safe_dict({
            "id":          int(s.get("id", 0)),
            "nom_client":  str(s.get("nom_client", "")),
            "type_client": str(s.get("type_client", "")),
        }))

    contacts_clean = []
    for _, ct in contacts_df.iterrows():
        contacts_clean.append(_safe_dict({
            "id":     int(ct["id"]),
            "prenom": str(ct.get("prenom", "")),
            "nom":    str(ct.get("nom", "")),
            "role":   str(ct.get("role", "")),
            "email":  str(ct.get("email", "")),
            "is_primary": bool(int(ct.get("is_primary", 0) or 0)),
        }))
    activities_clean = []
    for _, a in activities_df.iterrows():
        activities_clean.append(_safe_dict({
            "id":               int(a["id"]),
            "date":             str(a.get("date", "")),
            "notes":            str(a.get("notes", "") or ""),
            "type_interaction": str(a.get("type_interaction", "") or ""),
        }))

    return _safe_dict({
        "id":             int(c["id"]),
        "nom_client":     str(c["nom_client"]),
        "type_client":    str(c.get("type_client", "")),
        "region":         str(c.get("region", "")),
        "country":        str(c.get("country", "")),
        "tier":           str(c.get("tier", "Tier 2")),
        "kyc_status":     str(c.get("kyc_status", "En cours")),
        "parent_nom":     str(c.get("parent_nom", "") or ""),
        "primary_contact": primary_contact,
        "aum_funded":     funded_aum,
        "aum_active":     active_aum,
        "aum_total":      funded_aum + active_aum,
        "nb_active_deals": int(len(active)),
        "nb_activities":   int(len(activities_clean)),
        "dominant_statut": dominant_statut,
        "fonds_invested": fonds_invested,
        "created_year":   None,
        "network": {
            "root":         network.get("root"),
            "subsidiaries": subs_clean,
            "fund_links":   fund_links_clean,
        },
        "contacts":   contacts_clean,
        "activities": activities_clean,
    })


# ---------------------------------------------------------------------------
# /api/activities — recent activities
# ---------------------------------------------------------------------------

@app.get("/api/activities")
def api_activities():
    df = db.get_activities()
    if df.empty:
        return []
    out = []
    for _, a in df.iterrows():
        out.append(_safe_dict({
            "id":               int(a["id"]),
            "nom_client":       str(a.get("nom_client", "")),
            "date":             str(a.get("date", "")),
            "type_interaction": str(a.get("type_interaction", "") or ""),
            "notes":            str(a.get("notes", "") or ""),
        }))
    return out


# ---------------------------------------------------------------------------
# /api/sales — sales team metrics (Funded/Pipeline/Conversion per rep)
# ---------------------------------------------------------------------------

@app.get("/api/sales")
def api_sales():
    team_df = db.get_sales_team()
    metrics_df = db.get_sales_metrics()

    by_owner = {}
    if not metrics_df.empty:
        for _, m in metrics_df.iterrows():
            by_owner[str(m["Commercial"])] = m

    out = []
    for _, t in team_df.iterrows():
        nom = str(t["nom"])
        m = by_owner.get(nom, None)
        funded_aum  = float(m["AUM_Finance"]) if m is not None else 0.0
        pipeline    = float(m["Pipeline_Actif"]) if m is not None else 0.0
        nb_actifs   = int(m["Actifs"]) if m is not None else 0
        nb_funded   = int(m["Funded"]) if m is not None else 0
        nb_perdus   = int(m["Perdus"]) if m is not None else 0
        nb_total    = nb_funded + nb_perdus
        conversion  = round(nb_funded / nb_total * 100, 0) if nb_total > 0 else 0
        out.append(_safe_dict({
            "id":          int(t["id"]),
            "nom":         nom,
            "marche":      str(t.get("marche", "")),
            "funded_aum":  funded_aum,
            "pipeline_aum": pipeline,
            "nb_actifs":   nb_actifs,
            "nb_funded":   nb_funded,
            "conversion":  int(conversion),
        }))
    out.sort(key=lambda s: s["funded_aum"], reverse=True)
    return out


# ---------------------------------------------------------------------------
# /api/regions — AUM Funded by region
# ---------------------------------------------------------------------------

@app.get("/api/regions")
def api_regions():
    region_aum = db.get_aum_by_region()
    out = [{"region": k, "aum": float(v)} for k, v in region_aum.items()]
    out.sort(key=lambda r: r["aum"], reverse=True)
    return out


# ---------------------------------------------------------------------------
# /api/whitespace — top 5 clients × all funds
# ---------------------------------------------------------------------------

@app.get("/api/whitespace")
def api_whitespace():
    matrix = db.get_whitespace_matrix()
    if matrix is None or matrix.empty:
        return {"clients": [], "fonds": [], "values": []}
    totals = matrix.fillna(0).sum(axis=1).sort_values(ascending=False)
    top_clients = totals.head(5).index.tolist()
    sub = matrix.loc[top_clients]
    fonds = sub.columns.tolist()
    values = []
    for _, row in sub.iterrows():
        values.append([
            0.0 if (v is None or (isinstance(v, float) and np.isnan(v)))
            else float(v) for v in row
        ])
    return {"clients": top_clients, "fonds": fonds, "values": values}


# ---------------------------------------------------------------------------
# /api/next-actions — upcoming pipeline actions (30 days)
# ---------------------------------------------------------------------------

@app.get("/api/next-actions")
def api_next_actions():
    today = date.today()
    horizon = today + timedelta(days=30)
    overdue_cutoff = today - timedelta(days=60)
    df = db.get_pipeline_with_clients()
    if df.empty:
        return []
    df = df[df["statut"].isin(
        ["Prospect","Initial Pitch","Due Diligence","Soft Commit"])].copy()
    out = []
    for _, r in df.iterrows():
        nad = r.get("next_action_date")
        if not isinstance(nad, date):
            continue
        if not (overdue_cutoff <= nad <= horizon):
            continue
        out.append(_safe_dict({
            "id":               int(r["id"]),
            "nom_client":       str(r["nom_client"]),
            "fonds":            str(r.get("fonds", "")),
            "statut":           str(r.get("statut", "")),
            "next_action_date": nad.isoformat(),
            "sales_owner":      str(r.get("sales_owner", "")),
        }))
    out.sort(key=lambda x: x["next_action_date"])
    return out


# ---------------------------------------------------------------------------
# /api/funds-aggregate — Funded AUM aggregated by fund name
# ---------------------------------------------------------------------------

@app.get("/api/funds-aggregate")
def api_funds_aggregate():
    df = db.get_pipeline_with_clients()
    if df.empty:
        return {}
    df_f = df[df["statut"] == "Funded"]
    agg = df_f.groupby("fonds")["funded_aum"].sum().to_dict()
    return {str(k): float(v) for k, v in agg.items()}


# ---------------------------------------------------------------------------
# /api/monthly-aum — synthetic monthly AUM evolution (last 12 months)
# ---------------------------------------------------------------------------

@app.get("/api/monthly-aum")
def api_monthly_aum():
    df = db.get_historical_aum(days_back=365) if hasattr(db, "get_historical_aum") else None
    if df is None or df.empty:
        return []
    df = df.copy()
    df["month"] = pd.to_datetime(df["date"]).dt.to_period("M")
    agg = df.groupby("month")["funded_aum"].max().reset_index()
    fr_months = ["jan","fév","mar","avr","mai","juin",
                 "juil","août","sept","oct","nov","déc"]
    out = []
    for _, r in agg.tail(12).iterrows():
        period = r["month"]
        out.append({
            "label":  fr_months[period.month - 1],
            "month":  str(period),
            "aum":    float(r["funded_aum"]),
        })
    return out


# ---------------------------------------------------------------------------
# /api/mini-stats — bottom row mini-stat cards
# ---------------------------------------------------------------------------

@app.get("/api/mini-stats")
def api_mini_stats():
    conn = db.get_connection()
    c = conn.cursor()
    c.execute("SELECT COUNT(*) FROM clients"); nb_clients = int(c.fetchone()[0])
    c.execute("SELECT COUNT(*) FROM contacts"); nb_contacts = int(c.fetchone()[0])
    cutoff_30d = (date.today() - timedelta(days=30)).isoformat()
    c.execute("SELECT COUNT(*) FROM activites WHERE date >= ?", (cutoff_30d,))
    nb_activities_30d = int(c.fetchone()[0])
    today_str = date.today().isoformat()
    c.execute(
        "SELECT COUNT(*) FROM pipeline"
        " WHERE next_action_date < ?"
        " AND statut NOT IN ('Lost','Funded','Redeemed')",
        (today_str,))
    nb_overdue = int(c.fetchone()[0])
    c.execute("SELECT COUNT(DISTINCT fonds) FROM pipeline WHERE fonds != ''")
    nb_fonds = int(c.fetchone()[0])
    conn.close()
    return {
        "nb_clients":         nb_clients,
        "nb_contacts":        nb_contacts,
        "nb_activities_30d":  nb_activities_30d,
        "nb_overdue":         nb_overdue,
        "nb_fonds":           nb_fonds,
    }


# ---------------------------------------------------------------------------
# /api/overview — single aggregate endpoint that powers the whole dashboard
# ---------------------------------------------------------------------------

@app.get("/api/overview")
def api_overview():
    return {
        "kpis":              api_kpis(),
        "top_deals":         api_deals()[:8],
        "all_deals":         api_deals(),
        "statut_counts":     api_kpis()["statut_repartition"],
        "recent_activities": api_activities(),
        "whitespace":        api_whitespace(),
        "regions":           api_regions(),
        "next_actions":      api_next_actions(),
        "sales_team":        api_sales(),
        "mini_stats":        api_mini_stats(),
        "funds_aggregate":   api_funds_aggregate(),
        "monthly_aum":       api_monthly_aum(),
    }


# ---------------------------------------------------------------------------
# Local launcher
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
