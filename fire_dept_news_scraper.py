"""
Fire Department News Scraper
============================
Queries Google News RSS for fire department news across cities in California.
Outputs a timestamped Excel file + CSV ready to import into Google Sheets.

Usage:
    pip install requests pandas openpyxl
    python fire_dept_news_scraper.py

Output:
    fire_dept_news_YYYY-MM-DD.xlsx
    fire_dept_news_YYYY-MM-DD.csv
"""

import requests
import xml.etree.ElementTree as ET
import pandas as pd
import time
import re
from datetime import datetime
from urllib.parse import quote

# ---------------------------------------------------------------------------
# CITY LIST
# ---------------------------------------------------------------------------
CITIES = {
    "California": [
        "Los Angeles", "San Diego", "San Jose", "San Francisco", "Sacramento",
        "Fresno", "Long Beach", "Oakland", "Bakersfield", "Anaheim",
        "Santa Ana", "Riverside", "Stockton", "Chula Vista", "Irvine",
        "Fremont", "San Bernardino", "Modesto", "Fontana", "Oxnard",
        "Moreno Valley", "Glendale", "Huntington Beach", "Santa Clarita",
        "Garden Grove", "Oceanside", "Rancho Cucamonga", "Santa Rosa",
        "Ontario", "Elk Grove",
    ],
}

# ---------------------------------------------------------------------------
# SEARCH QUERIES  (each city will be substituted for {city})
# ---------------------------------------------------------------------------
QUERY_TEMPLATES = [
    '"{city}" fire department',
    '"{city}" fire department budget',
    '"{city}" fire department staffing',
    '"{city}" fire department chief',
    '"{city}" wildfire emergency',
    '"{city}" fire department incident',
]

# ---------------------------------------------------------------------------
# CATEGORY KEYWORDS
# ---------------------------------------------------------------------------
CATEGORIES = {
    "Budget/Funding":        ["budget", "fund", "funding", "cost", "tax", "levy", "million", "grant", "cuts", "shortfall"],
    "Staffing/Labor":        ["hire", "hiring", "staff", "staffing", "recruit", "union", "salary", "wage", "retire", "pension", "layoff", "overtime"],
    "Leadership/Personnel":  ["chief", "appoint", "resign", "retire", "fired", "promot", "director", "captain", "officer"],
    "Incident/Disaster":     ["wildfire", "fire", "blaze", "disaster", "emergency", "rescue", "response", "fatal", "evacuat", "explosion", "accident", "injury", "death"],
    "Equipment/Stations":    ["truck", "engine", "equipment", "apparatus", "station", "vehicle", "aerial", "ladder", "purchase"],
    "Policy/Legislation":    ["policy", "ordinance", "law", "legislation", "reform", "measure", "vote", "council", "ballot", "proposition"],
}

def categorize(headline: str, snippet: str) -> str:
    text = (headline + " " + snippet).lower()
    for category, keywords in CATEGORIES.items():
        if any(kw in text for kw in keywords):
            return category
    return "General"

# ---------------------------------------------------------------------------
# GOOGLE NEWS RSS FETCH
# ---------------------------------------------------------------------------
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
}

def fetch_rss(query: str, max_results: int = 10) -> list[dict]:
    url = (
        f"https://news.google.com/rss/search"
        f"?q={quote(query)}&hl=en-US&gl=US&ceid=US:en"
    )
    try:
        resp = requests.get(url, headers=HEADERS, timeout=12)
        resp.raise_for_status()
    except Exception as e:
        print(f"  [WARN] RSS fetch failed for '{query}': {e}")
        return []

    try:
        root = ET.fromstring(resp.content)
    except ET.ParseError as e:
        print(f"  [WARN] XML parse error for '{query}': {e}")
        return []

    items = []
    for item in root.findall(".//item")[:max_results]:
        title   = (item.findtext("title")       or "").strip()
        link    = (item.findtext("link")        or "").strip()
        pub     = (item.findtext("pubDate")     or "").strip()
        desc    = (item.findtext("description") or "").strip()
        source  = (item.findtext("source")      or "").strip()

        # Clean HTML tags from description
        desc_clean = re.sub(r"<[^>]+>", "", desc).strip()

        # Parse date
        try:
            pub_dt = datetime.strptime(pub, "%a, %d %b %Y %H:%M:%S %Z")
            pub_str = pub_dt.strftime("%Y-%m-%d")
        except Exception:
            pub_str = pub[:10] if len(pub) >= 10 else pub

        items.append({
            "headline": title,
            "url":      link,
            "date":     pub_str,
            "source":   source,
            "snippet":  desc_clean[:300],
        })
    return items

# ---------------------------------------------------------------------------
# HTML EXPORT
# ---------------------------------------------------------------------------

CATEGORY_COLORS = {
    "Budget/Funding":       "#3b82f6",
    "Staffing/Labor":       "#8b5cf6",
    "Leadership/Personnel": "#f59e0b",
    "Incident/Disaster":    "#ef4444",
    "Equipment/Stations":   "#10b981",
    "Policy/Legislation":   "#6366f1",
    "General":              "#6b7280",
}

def build_html(df: "pd.DataFrame", path: str, today: str) -> None:
    import json

    records = df.to_dict(orient="records")
    data_json = json.dumps(records, ensure_ascii=False)

    cities    = sorted(df["City"].unique().tolist())
    cats      = sorted(df["Category"].unique().tolist())
    cat_colors_json = json.dumps(CATEGORY_COLORS)

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>CA Fire Dept News — {today}</title>
<style>
  *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
         background: #f1f5f9; color: #1e293b; min-height: 100vh; }}
  header {{ background: #1e3a5f; color: #fff; padding: 20px 32px; }}
  header h1 {{ font-size: 1.4rem; font-weight: 700; letter-spacing: -.3px; }}
  header p  {{ font-size: .85rem; opacity: .75; margin-top: 4px; }}
  .controls {{ display: flex; flex-wrap: wrap; gap: 12px; padding: 18px 32px;
               background: #fff; border-bottom: 1px solid #e2e8f0; align-items: center; }}
  .controls input, .controls select {{
    border: 1px solid #cbd5e1; border-radius: 8px; padding: 8px 12px;
    font-size: .875rem; outline: none; background: #f8fafc; color: #1e293b; }}
  .controls input {{ flex: 1; min-width: 220px; }}
  .controls input:focus, .controls select:focus {{ border-color: #3b82f6; background: #fff; }}
  #count {{ margin-left: auto; font-size: .82rem; color: #64748b; white-space: nowrap; }}
  .table-wrap {{ overflow-x: auto; padding: 20px 32px 40px; }}
  table {{ width: 100%; border-collapse: collapse; background: #fff;
           border-radius: 10px; overflow: hidden;
           box-shadow: 0 1px 3px rgba(0,0,0,.08); font-size: .875rem; }}
  thead tr {{ background: #1e3a5f; color: #fff; }}
  th {{ padding: 11px 14px; text-align: left; font-weight: 600;
        cursor: pointer; user-select: none; white-space: nowrap; }}
  th:hover {{ background: #2d4f7c; }}
  th .arrow {{ opacity: .45; margin-left: 4px; font-size: .7rem; }}
  th.asc  .arrow::after {{ content: "▲"; opacity: 1; }}
  th.desc .arrow::after {{ content: "▼"; opacity: 1; }}
  th:not(.asc):not(.desc) .arrow::after {{ content: "⇅"; }}
  tbody tr {{ border-bottom: 1px solid #f1f5f9; transition: background .1s; }}
  tbody tr:hover {{ background: #f8fafc; }}
  td {{ padding: 10px 14px; vertical-align: top; }}
  td.date   {{ white-space: nowrap; color: #64748b; font-size: .8rem; }}
  td.city   {{ white-space: nowrap; font-weight: 500; }}
  td.source {{ white-space: nowrap; color: #64748b; font-size: .8rem; }}
  td.snippet {{ color: #475569; font-size: .82rem; max-width: 340px; line-height: 1.45; }}
  a.headline {{ color: #1e3a5f; font-weight: 500; text-decoration: none; line-height: 1.4; }}
  a.headline:hover {{ text-decoration: underline; color: #3b82f6; }}
  .badge {{ display: inline-block; padding: 2px 9px; border-radius: 99px;
            font-size: .75rem; font-weight: 600; color: #fff; white-space: nowrap; }}
  .empty {{ text-align: center; padding: 60px; color: #94a3b8; font-size: .95rem; }}
</style>
</head>
<body>

<header>
  <h1>🔥 California Fire Department News</h1>
  <p>Generated {today} &nbsp;·&nbsp; Google News RSS &nbsp;·&nbsp; {len(df)} articles across {df["City"].nunique()} cities</p>
</header>

<div class="controls">
  <input id="search" type="text" placeholder="Search headlines, cities, sources…" oninput="render()">
  <select id="filterCity" onchange="render()">
    <option value="">All Cities</option>
    {''.join(f'<option value="{c}">{c}</option>' for c in cities)}
  </select>
  <select id="filterCat" onchange="render()">
    <option value="">All Categories</option>
    {''.join(f'<option value="{c}">{c}</option>' for c in cats)}
  </select>
  <span id="count"></span>
</div>

<div class="table-wrap">
  <table>
    <thead>
      <tr>
        <th onclick="sortBy('Date')"    data-col="Date">    Date    <span class="arrow"></span></th>
        <th onclick="sortBy('City')"    data-col="City">    City    <span class="arrow"></span></th>
        <th onclick="sortBy('Category')"data-col="Category">Category<span class="arrow"></span></th>
        <th>Headline</th>
        <th onclick="sortBy('Source')"  data-col="Source">  Source  <span class="arrow"></span></th>
        <th>Snippet</th>
      </tr>
    </thead>
    <tbody id="tbody"></tbody>
  </table>
</div>

<script>
const DATA   = {data_json};
const COLORS = {cat_colors_json};
let sortCol  = "Date";
let sortDir  = -1;  // -1 = desc

function sortBy(col) {{
  if (sortCol === col) sortDir *= -1;
  else {{ sortCol = col; sortDir = -1; }}
  document.querySelectorAll("th[data-col]").forEach(th => {{
    th.classList.remove("asc","desc");
    if (th.dataset.col === col) th.classList.add(sortDir === 1 ? "asc" : "desc");
  }});
  render();
}}

function render() {{
  const q    = document.getElementById("search").value.toLowerCase();
  const city = document.getElementById("filterCity").value;
  const cat  = document.getElementById("filterCat").value;

  let rows = DATA.filter(r => {{
    if (city && r.City     !== city) return false;
    if (cat  && r.Category !== cat)  return false;
    if (q) {{
      const haystack = (r.Headline + r.City + r.Source + r.Snippet + r.Category).toLowerCase();
      if (!haystack.includes(q)) return false;
    }}
    return true;
  }});

  rows.sort((a,b) => {{
    const av = (a[sortCol] || "").toString();
    const bv = (b[sortCol] || "").toString();
    return av < bv ? -sortDir : av > bv ? sortDir : 0;
  }});

  document.getElementById("count").textContent =
    rows.length + " article" + (rows.length !== 1 ? "s" : "");

  const tbody = document.getElementById("tbody");
  if (!rows.length) {{
    tbody.innerHTML = '<tr><td colspan="6" class="empty">No articles match your filters.</td></tr>';
    return;
  }}

  tbody.innerHTML = rows.map(r => {{
    const color = COLORS[r.Category] || "#6b7280";
    return `<tr>
      <td class="date">${{r.Date}}</td>
      <td class="city">${{esc(r.City)}}</td>
      <td><span class="badge" style="background:${{color}}">${{esc(r.Category)}}</span></td>
      <td><a class="headline" href="${{esc(r.URL)}}" target="_blank" rel="noopener">${{esc(r.Headline)}}</a></td>
      <td class="source">${{esc(r.Source)}}</td>
      <td class="snippet">${{esc(r.Snippet)}}</td>
    </tr>`;
  }}).join("");
}}

function esc(s) {{
  return String(s||"")
    .replace(/&/g,"&amp;").replace(/</g,"&lt;")
    .replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}}

// Initial sort header state
document.querySelector('th[data-col="Date"]').classList.add("desc");
render();
</script>
</body>
</html>"""

    with open(path, "w", encoding="utf-8") as f:
        f.write(html)


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------
def main():
    all_rows   = []
    seen_urls  = set()
    today      = datetime.today().strftime("%Y-%m-%d")

    total_cities = sum(len(v) for v in CITIES.values())
    processed    = 0

    for state, cities in CITIES.items():
        print(f"\n{'='*60}")
        print(f"  STATE: {state}  ({len(cities)} cities)")
        print(f"{'='*60}")

        for city in cities:
            processed += 1
            print(f"  [{processed}/{total_cities}] {city}, {state}")

            for template in QUERY_TEMPLATES:
                query = template.format(city=city)
                results = fetch_rss(query)

                for r in results:
                    url = r["url"]
                    if url in seen_urls:
                        continue
                    seen_urls.add(url)

                    category = categorize(r["headline"], r["snippet"])

                    all_rows.append({
                        "State":     state,
                        "City":      city,
                        "Category":  category,
                        "Date":      r["date"],
                        "Headline":  r["headline"],
                        "Source":    r["source"],
                        "Snippet":   r["snippet"],
                        "URL":       r["url"],
                    })

                time.sleep(0.4)   # polite rate limit between queries

    # -----------------------------------------------------------------------
    # BUILD DATAFRAME & EXPORT
    # -----------------------------------------------------------------------
    if not all_rows:
        print("\n[ERROR] No results collected. Check network access.")
        return

    df = pd.DataFrame(all_rows)
    df = df.sort_values(["State", "City", "Date"], ascending=[True, True, False])

    xlsx_path = f"fire_dept_news_{today}.xlsx"
    csv_path  = f"fire_dept_news_{today}.csv"

    # --- Excel with light formatting ---
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="News")

        ws = writer.sheets["News"]

        # Column widths
        col_widths = {
            "A": 14,   # State
            "B": 20,   # City
            "C": 22,   # Category
            "D": 12,   # Date
            "E": 60,   # Headline
            "F": 22,   # Source
            "G": 60,   # Snippet
            "H": 50,   # URL
        }
        for col, width in col_widths.items():
            ws.column_dimensions[col].width = width

        # Freeze header row
        ws.freeze_panes = "A2"

        # Auto-filter
        ws.auto_filter.ref = ws.dimensions

        # Summary sheet
        summary = df.groupby(["State", "City", "Category"]).size().reset_index(name="Article Count")
        summary.to_excel(writer, index=False, sheet_name="Summary")

    # --- CSV for Google Sheets import ---
    df.to_csv(csv_path, index=False)

    # --- Searchable HTML page ---
    html_path = f"fire_dept_news_{today}.html"
    build_html(df, html_path, today)

    print(f"\n{'='*60}")
    print(f"  Done!  {len(df)} unique articles across {df['City'].nunique()} cities")
    print(f"  Excel: {xlsx_path}")
    print(f"  CSV:   {csv_path}  ← drag this into Google Sheets")
    print(f"  HTML:  {html_path}  ← open in any browser")
    print(f"{'='*60}")

if __name__ == "__main__":
    main()
