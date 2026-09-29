"""
APMI PMS scraper: monthly returns, AUM, and portfolio turnover (1M / 1Y)
Source: https://www.apmiindia.org (IA Performance Report + IA Turnover Report)

Install:   pip install requests beautifulsoup4 pandas openpyxl lxml
Run:       python apmi_scraper.py                      # latest month, default filters
           python apmi_scraper.py --inspect            # print the report forms' fields
           python apmi_scraper.py --months 2026-06 2026-07 2026-08   # specific months (best effort)

Output:    apmi_pms_data.xlsx  (sheets: Combined, Performance, Turnover)
"""

import argparse
import re
import sys
import time
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup

BASE = "https://www.apmiindia.org/apmi/"
PERF_URL = BASE + "welcomeiaperformance.htm?action=PMSmenu"
TURN_URL = BASE + "IATurnoverReportUtility.htm?action=loadIATurnoverPage"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (research script; contact: you@example.com)",
    "Accept-Language": "en-IN,en;q=0.9",
}
DELAY_SECONDS = 3  # be polite between requests


# ---------------------------------------------------------------- helpers
def clean_number(value):
    """'₹1,983.43' -> 1983.43, 'NA' -> None, '-1.03' -> -1.03"""
    if value is None:
        return None
    text = str(value).replace("₹", "").replace(",", "").strip()
    if text.upper() in {"", "NA", "N/A", "-", "--"}:
        return None
    try:
        return float(text)
    except ValueError:
        return text


def norm(text):
    return re.sub(r"\s+", " ", str(text)).strip().lower()


def find_table(html, required_headers, forbidden_headers=()):
    """Return the first <table> whose header row contains all required labels."""
    soup = BeautifulSoup(html, "lxml")
    for table in soup.find_all("table"):
        rows = table.find_all("tr")
        if not rows:
            continue
        # header = first row that has th cells, else first row
        header_row = next((r for r in rows if r.find("th")), rows[0])
        headers = [norm(c.get_text()) for c in header_row.find_all(["th", "td"])]
        if all(any(req in h for h in headers) for req in required_headers) and not any(
            any(bad in h for h in headers) for bad in forbidden_headers
        ):
            data = []
            for r in rows[rows.index(header_row) + 1 :]:
                cells = [c.get_text(" ", strip=True) for c in r.find_all("td")]
                if len(cells) == len(headers):
                    data.append(cells)
            return headers, data
    return None, None


def extract_as_on(html):
    m = re.search(r"As on\s*(\d{2}/\d{2}/\d{4})", BeautifulSoup(html, "lxml").get_text(" "))
    return m.group(1) if m else None


# ---------------------------------------------------------------- parsers
PERF_COLS = {
    "pms provider name": "Provider",
    "ia name": "Investment Approach",
    "aum (in inr cr.)": "AUM (INR Cr)",
    "1 month": "Return 1M (%)",
    "3 months": "Return 3M (%)",
    "6 months": "Return 6M (%)",
    "1 year": "Return 1Y (%)",
    "2 years": "Return 2Y (%)",
    "3 years": "Return 3Y (%)",
    "4 years": "Return 4Y (%)",
    "5 years": "Return 5Y (%)",
    "since inception": "Return Since Inception (%)",
}
TURN_COLS = {
    "pms provider name": "Provider",
    "ia name": "Investment Approach",
    "1 month": "Turnover 1M",
    "1 year": "Turnover 1Y",
}


def parse_report(html, required, forbidden, colmap):
    headers, rows = find_table(html, required, forbidden)
    if headers is None:
        raise RuntimeError(
            "Could not find the data table. The page layout may have changed; "
            "save the HTML and inspect it."
        )
    df = pd.DataFrame(rows, columns=headers)
    df = df.rename(columns={h: colmap[h] for h in headers if h in colmap})
    df = df[[c for c in colmap.values() if c in df.columns]]
    for col in df.columns:
        if col not in ("Provider", "Investment Approach"):
            df[col] = df[col].map(clean_number)
    return df


def parse_performance(html):
    return parse_report(html, ["aum", "ia name", "1 month"], [], PERF_COLS)


def parse_turnover(html):
    return parse_report(html, ["ia name", "1 month", "1 year"], ["aum"], TURN_COLS)


# ---------------------------------------------------------------- form handling (history)
def describe_forms(html, label):
    soup = BeautifulSoup(html, "lxml")
    print(f"\n=== Forms on {label} ===")
    for i, form in enumerate(soup.find_all("form")):
        print(f"\nForm #{i}: action={form.get('action')!r} method={form.get('method')!r}")
        for el in form.find_all(["input", "select", "textarea"]):
            name = el.get("name") or el.get("id")
            if el.name == "select":
                opts = [o.get_text(strip=True) for o in el.find_all("option")][:6]
                print(f"  select {name!r}: first options {opts}")
            else:
                print(f"  {el.name} type={el.get('type')!r} name={name!r} value={el.get('value')!r}")


def submit_for_month(session, page_url, html, month):
    """
    Best effort: re-submit the report form with a different 'As On Month-Year'.
    month is 'YYYY-MM'. Keeps all other fields at their defaults.
    If this doesn't work, run --inspect and adjust MONTH_FORMATS / field matching.
    """
    soup = BeautifulSoup(html, "lxml")
    form = next(
        (f for f in soup.find_all("form") if "month" in norm(f.get_text()) or f.find("select")),
        None,
    )
    if form is None:
        raise RuntimeError("No report form found; run with --inspect.")

    payload = {}
    for el in form.find_all(["input", "select"]):
        name = el.get("name")
        if not name or el.get("type") in ("submit", "button", "reset"):
            continue
        if el.name == "select":
            chosen = el.find("option", selected=True) or el.find("option")
            payload[name] = chosen.get("value", chosen.get_text(strip=True)) if chosen else ""
        elif el.get("type") in ("checkbox", "radio"):
            if el.has_attr("checked"):
                payload[name] = el.get("value", "on")
        else:
            payload[name] = el.get("value", "")

    yyyy, mm = month.split("-")
    month_field = next(
        (k for k in payload if re.search(r"month|date|year|ason", k, re.I)), None
    )
    if month_field is None:
        raise RuntimeError("Couldn't identify the month field; run with --inspect.")
    # Guess the format from the current value (e.g. '08-2026', 'Aug-2026', '2026-08')
    current = str(payload[month_field])
    if re.fullmatch(r"\d{2}-\d{4}", current):
        payload[month_field] = f"{mm}-{yyyy}"
    elif re.fullmatch(r"\d{2}/\d{4}", current):
        payload[month_field] = f"{mm}/{yyyy}"
    elif re.fullmatch(r"[A-Za-z]{3}-\d{4}", current):
        payload[month_field] = pd.Timestamp(f"{yyyy}-{mm}-01").strftime("%b-%Y")
    else:
        payload[month_field] = f"{yyyy}-{mm}"

    action = urljoin(page_url, form.get("action") or page_url)
    method = (form.get("method") or "post").lower()
    resp = session.request(method, action, data=payload if method == "post" else None,
                           params=payload if method == "get" else None, timeout=60)
    resp.raise_for_status()
    return resp.text


# ---------------------------------------------------------------- main
def fetch(session, url):
    resp = session.get(url, timeout=60)
    resp.raise_for_status()
    return resp.text


def merge(perf, turn):
    # Some providers list the same IA name twice; number duplicates so rows pair up correctly
    for df in (perf, turn):
        df["_key"] = df["Provider"].map(norm) + "|" + df["Investment Approach"].map(norm)
        df["_dup"] = df.groupby("_key").cumcount()
    combined = perf.merge(
        turn.drop(columns=["Provider", "Investment Approach"]),
        on=["_key", "_dup"], how="outer", indicator=True,
    )
    # fill names for turnover-only rows
    missing = combined["Provider"].isna()
    if missing.any():
        names = turn.set_index(["_key", "_dup"])[["Provider", "Investment Approach"]]
        idx = list(zip(combined.loc[missing, "_key"], combined.loc[missing, "_dup"]))
        combined.loc[missing, ["Provider", "Investment Approach"]] = names.loc[idx].values
    combined["Matched In"] = combined["_merge"].map(
        {"both": "Both reports", "left_only": "Performance only", "right_only": "Turnover only"}
    )
    for df in (perf, turn):
        df.drop(columns=["_key", "_dup"], inplace=True)
    return combined.drop(columns=["_key", "_dup", "_merge"])


def scrape_one(session, month=None):
    perf_html = fetch(session, PERF_URL)
    time.sleep(DELAY_SECONDS)
    turn_html = fetch(session, TURN_URL)
    if month:
        time.sleep(DELAY_SECONDS)
        perf_html = submit_for_month(session, PERF_URL, perf_html, month)
        time.sleep(DELAY_SECONDS)
        turn_html = submit_for_month(session, TURN_URL, turn_html, month)

    perf, turn = parse_performance(perf_html), parse_turnover(turn_html)
    perf_date, turn_date = extract_as_on(perf_html), extract_as_on(turn_html)
    if perf_date != turn_date:
        print(f"  Warning: performance is as on {perf_date}, turnover as on {turn_date}")
    combined = merge(perf, turn)
    for df, d in ((perf, perf_date), (turn, turn_date), (combined, perf_date)):
        df.insert(0, "As On", d)
    print(f"  {perf_date}: {len(perf)} performance rows, {len(turn)} turnover rows")
    return combined, perf, turn


def write_excel(path, sheets):
    with pd.ExcelWriter(path, engine="openpyxl") as xl:
        for name, df in sheets.items():
            df.to_excel(xl, sheet_name=name, index=False)
            ws = xl.sheets[name]
            ws.freeze_panes = "D2" if name != "Turnover" else "C2"
            ws.auto_filter.ref = ws.dimensions
            for col in ws.columns:
                width = max(len(str(c.value or "")) for c in col[:200])
                ws.column_dimensions[col[0].column_letter].width = min(max(width + 2, 10), 60)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--months", nargs="*", help="YYYY-MM months to fetch (best effort)")
    ap.add_argument("--inspect", action="store_true", help="print form fields and exit")
    ap.add_argument("--out", default="apmi_pms_data.xlsx")
    ap.add_argument("--json", help="also write the Combined data as JSON for the website, e.g. docs/data.json")
    args = ap.parse_args()

    session = requests.Session()
    session.headers.update(HEADERS)

    if args.inspect:
        describe_forms(fetch(session, PERF_URL), "IA Performance")
        describe_forms(fetch(session, TURN_URL), "IA Turnover")
        return

    months = args.months or [None]
    combined_all, perf_all, turn_all = [], [], []
    for m in months:
        print(f"Fetching {m or 'latest month'} ...")
        try:
            c, p, t = scrape_one(session, m)
        except Exception as e:  # keep going for other months
            print(f"  Failed for {m}: {e}", file=sys.stderr)
            continue
        combined_all.append(c); perf_all.append(p); turn_all.append(t)

    if not combined_all:
        sys.exit("No data collected.")
    write_excel(args.out, {
        "Combined": pd.concat(combined_all, ignore_index=True),
        "Performance": pd.concat(perf_all, ignore_index=True),
        "Turnover": pd.concat(turn_all, ignore_index=True),
    })
    print(f"Saved {args.out}")

    if args.json:
        import json
        from datetime import datetime, timezone
        combined = pd.concat(combined_all, ignore_index=True)
        combined = combined.astype(object).where(pd.notna(combined), None)
        payload = {
            "as_on": combined["As On"].iloc[0],
            "updated_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC"),
            "source": "Association of Portfolio Managers in India (apmiindia.org)",
            "rows": combined.to_dict(orient="records"),
        }
        with open(args.json, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False)
        print(f"Saved {args.json} ({len(payload['rows'])} strategies)")


if __name__ == "__main__":
    main()
