"""
California Fire Department News — Local Web App
================================================
Install:  pip install fastapi "uvicorn[standard]" requests pandas openpyxl
Run:      python app.py
Open:     http://localhost:8000
"""

from fastapi import FastAPI, Query, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, StreamingResponse
from pydantic import BaseModel
import uvicorn
import os
import sqlite3
import json
import io
import re
import time
import threading
from datetime import datetime
from urllib.parse import quote
from pathlib import Path
from typing import Optional
import requests as http_requests
import pandas as pd

# ---------------------------------------------------------------------------
# PATHS
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).parent
DB_PATH  = BASE_DIR / "data" / "articles.db"
CFG_PATH = BASE_DIR / "data" / "cities.json"

# ---------------------------------------------------------------------------
# DATABASE
# ---------------------------------------------------------------------------
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    with get_db() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS articles (
                id       INTEGER PRIMARY KEY AUTOINCREMENT,
                state    TEXT,
                city     TEXT,
                category TEXT,
                date     TEXT,
                headline TEXT,
                source   TEXT,
                snippet  TEXT,
                url      TEXT UNIQUE,
                scraped_at TEXT
            )
        """)
        conn.commit()

# ---------------------------------------------------------------------------
# SCRAPER LOGIC
# ---------------------------------------------------------------------------
QUERY_TEMPLATES = [
    '"{city}" fire department',
    '"{city}" fire department budget',
    '"{city}" fire department staffing',
    '"{city}" fire department chief',
    '"{city}" wildfire emergency',
    '"{city}" fire department incident',
]

CATEGORIES = {
    "Budget/Funding":        ["budget","fund","funding","cost","tax","levy","million","grant","cuts","shortfall"],
    "Staffing/Labor":        ["hire","hiring","staff","staffing","recruit","union","salary","wage","retire","pension","layoff","overtime"],
    "Leadership/Personnel":  ["chief","appoint","resign","retire","fired","promot","director","captain","officer"],
    "Incident/Disaster":     ["wildfire","fire","blaze","disaster","emergency","rescue","response","fatal","evacuat","explosion","accident","injury","death"],
    "Equipment/Stations":    ["truck","engine","equipment","apparatus","station","vehicle","aerial","ladder","purchase"],
    "Policy/Legislation":    ["policy","ordinance","law","legislation","reform","measure","vote","council","ballot"],
}

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
}

def categorize(headline: str, snippet: str) -> str:
    import xml.etree.ElementTree as ET
    text = (headline + " " + snippet).lower()
    for cat, kws in CATEGORIES.items():
        if any(k in text for k in kws):
            return cat
    return "General"

def fetch_rss(query: str, max_results: int = 10) -> list:
    import xml.etree.ElementTree as ET
    url = f"https://news.google.com/rss/search?q={quote(query)}&hl=en-US&gl=US&ceid=US:en"
    try:
        resp = http_requests.get(url, headers=HEADERS, timeout=12)
        resp.raise_for_status()
        root = ET.fromstring(resp.content)
    except Exception:
        return []

    items = []
    for item in root.findall(".//item")[:max_results]:
        title  = (item.findtext("title")       or "").strip()
        link   = (item.findtext("link")        or "").strip()
        pub    = (item.findtext("pubDate")     or "").strip()
        desc   = (item.findtext("description") or "").strip()
        source = (item.findtext("source")      or "").strip()
        desc_clean = re.sub(r"<[^>]+>", "", desc).strip()
        try:
            pub_str = datetime.strptime(pub, "%a, %d %b %Y %H:%M:%S %Z").strftime("%Y-%m-%d")
        except Exception:
            pub_str = pub[:10]
        items.append({"headline": title, "url": link, "date": pub_str,
                       "source": source, "snippet": desc_clean[:300]})
    return items

# ---------------------------------------------------------------------------
# SCRAPE STATE  (shared between thread and API)
# ---------------------------------------------------------------------------
scrape_state = {
    "running":  False,
    "progress": 0,
    "total":    0,
    "current":  "",
    "new_count": 0,
    "error":    None,
    "last_run": None,
}
scrape_lock = threading.Lock()

def run_scrape():
    with scrape_lock:
        scrape_state["running"]   = True
        scrape_state["progress"]  = 0
        scrape_state["new_count"] = 0
        scrape_state["error"]     = None

    try:
        with open(CFG_PATH) as f:
            cities_cfg = json.load(f)

        pairs = [(state, city) for state, cities in cities_cfg.items() for city in cities]
        total = len(pairs) * len(QUERY_TEMPLATES)

        with scrape_lock:
            scrape_state["total"] = total

        done = 0
        new_count = 0
        scraped_at = datetime.utcnow().isoformat()

        conn = get_db()
        try:
            for state, city in pairs:
                with scrape_lock:
                    scrape_state["current"] = f"{city}, {state}"

                for template in QUERY_TEMPLATES:
                    query = template.format(city=city)
                    results = fetch_rss(query)

                    for r in results:
                        cat = categorize(r["headline"], r["snippet"])
                        try:
                            conn.execute(
                                """INSERT OR IGNORE INTO articles
                                   (state, city, category, date, headline, source, snippet, url, scraped_at)
                                   VALUES (?,?,?,?,?,?,?,?,?)""",
                                (state, city, cat, r["date"], r["headline"],
                                 r["source"], r["snippet"], r["url"], scraped_at)
                            )
                            if conn.execute("SELECT changes()").fetchone()[0]:
                                new_count += 1
                        except Exception:
                            pass

                    conn.commit()
                    done += 1
                    time.sleep(0.4)

                    with scrape_lock:
                        scrape_state["progress"]  = done
                        scrape_state["new_count"] = new_count
        finally:
            conn.close()

    except Exception as e:
        with scrape_lock:
            scrape_state["error"] = str(e)
    finally:
        with scrape_lock:
            scrape_state["running"]   = False
            scrape_state["last_run"]  = datetime.now().strftime("%Y-%m-%d %H:%M")
            scrape_state["current"]   = ""

# ---------------------------------------------------------------------------
# FASTAPI APP
# ---------------------------------------------------------------------------
app = FastAPI(title="Fire Dept News")

# ---------------------------------------------------------------------------
# API ROUTES
# ---------------------------------------------------------------------------

@app.get("/api/articles")
def get_articles(
    q:        Optional[str] = Query(None),
    city:     Optional[str] = Query(None),
    category: Optional[str] = Query(None),
    limit:    int = Query(500, le=2000),
    offset:   int = Query(0),
):
    sql    = "SELECT * FROM articles WHERE 1=1"
    params = []

    if city:
        sql += " AND city = ?"
        params.append(city)
    if category:
        sql += " AND category = ?"
        params.append(category)
    if q:
        sql += " AND (headline LIKE ? OR snippet LIKE ? OR source LIKE ?)"
        like = f"%{q}%"
        params += [like, like, like]

    sql += " ORDER BY date DESC, id DESC LIMIT ? OFFSET ?"
    params += [limit, offset]

    count_params = params[:-2]  # strip limit/offset

    with get_db() as conn:
        # Build a separate COUNT query from the same WHERE clause
        where_clause = sql[sql.index("WHERE"):sql.index("ORDER")].strip()
        count_sql    = f"SELECT COUNT(*) FROM articles {where_clause}"
        total        = conn.execute(count_sql, count_params).fetchone()[0]
        rows         = conn.execute(sql, params).fetchall()

    return {
        "total": total,
        "items": [dict(r) for r in rows],
    }


@app.get("/api/stats")
def get_stats():
    with get_db() as conn:
        total     = conn.execute("SELECT COUNT(*) FROM articles").fetchone()[0]
        cities    = conn.execute("SELECT COUNT(DISTINCT city) FROM articles").fetchone()[0]
        by_cat    = conn.execute(
            "SELECT category, COUNT(*) as n FROM articles GROUP BY category ORDER BY n DESC"
        ).fetchall()
        city_list = conn.execute(
            "SELECT DISTINCT city FROM articles ORDER BY city"
        ).fetchall()
        cat_list  = conn.execute(
            "SELECT DISTINCT category FROM articles ORDER BY category"
        ).fetchall()

    return {
        "total":     total,
        "cities":    cities,
        "by_cat":    [dict(r) for r in by_cat],
        "city_list": [r[0] for r in city_list],
        "cat_list":  [r[0] for r in cat_list],
    }


@app.post("/api/scrape/start")
def start_scrape():
    with scrape_lock:
        if scrape_state["running"]:
            raise HTTPException(status_code=409, detail="Scrape already running")
    t = threading.Thread(target=run_scrape, daemon=True)
    t.start()
    return {"status": "started"}


@app.get("/api/scrape/status")
def scrape_status():
    with scrape_lock:
        return dict(scrape_state)


# Cities management
@app.get("/api/cities")
def get_cities():
    with open(CFG_PATH) as f:
        return json.load(f)


class CitiesPayload(BaseModel):
    cities: dict  # {"California": ["Los Angeles", ...], ...}

@app.put("/api/cities")
def update_cities(payload: CitiesPayload):
    with open(CFG_PATH, "w") as f:
        json.dump(payload.cities, f, indent=2)
    return {"status": "saved"}


# Export
@app.get("/api/export/csv")
def export_csv(
    q:        Optional[str] = Query(None),
    city:     Optional[str] = Query(None),
    category: Optional[str] = Query(None),
):
    data = get_articles(q=q, city=city, category=category, limit=10000, offset=0)
    df   = pd.DataFrame(data["items"])
    if df.empty:
        raise HTTPException(404, "No data to export")
    buf = io.StringIO()
    df[["state","city","category","date","headline","source","snippet","url"]].to_csv(buf, index=False)
    buf.seek(0)
    today = datetime.today().strftime("%Y-%m-%d")
    return StreamingResponse(
        iter([buf.getvalue()]),
        media_type="text/csv",
        headers={"Content-Disposition": f"attachment; filename=fire_dept_news_{today}.csv"},
    )


@app.get("/api/export/xlsx")
def export_xlsx(
    q:        Optional[str] = Query(None),
    city:     Optional[str] = Query(None),
    category: Optional[str] = Query(None),
):
    data = get_articles(q=q, city=city, category=category, limit=10000, offset=0)
    df   = pd.DataFrame(data["items"])
    if df.empty:
        raise HTTPException(404, "No data to export")

    df = df[["state","city","category","date","headline","source","snippet","url"]]
    df.columns = ["State","City","Category","Date","Headline","Source","Snippet","URL"]

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="News")
        ws = writer.sheets["News"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for col, width in zip("ABCDEFGH", [14,20,22,12,60,22,60,50]):
            ws.column_dimensions[col].width = width
        summary = df.groupby(["State","City","Category"]).size().reset_index(name="Count")
        summary.to_excel(writer, index=False, sheet_name="Summary")
    buf.seek(0)

    today = datetime.today().strftime("%Y-%m-%d")
    return StreamingResponse(
        iter([buf.getvalue()]),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename=fire_dept_news_{today}.xlsx"},
    )


# ---------------------------------------------------------------------------
# STATIC FILES  (must be last so API routes take priority)
# ---------------------------------------------------------------------------
app.mount("/", StaticFiles(directory=str(BASE_DIR / "static"), html=True), name="static")


# ---------------------------------------------------------------------------
# ENTRY POINT
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    print("Initializing database…")
    init_db()
    port = int(os.environ.get("PORT", "8000"))
    print(f"Starting server at http://localhost:{port}")
    uvicorn.run(app, host="0.0.0.0", port=port, reload=False)
