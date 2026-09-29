"""
APMI PMS scraper: returns, AUM and portfolio turnover (1M / 1Y)
for every category (Equity, Debt, Hybrid, Multi Asset) and service type
(Discretionary, Non Discretionary).

Source: https://www.apmiindia.org (IA Performance Report + IA Turnover Report)

Install:   pip install requests beautifulsoup4 pandas openpyxl lxml
Run:       python apmi_scraper.py --json docs/data.json
           python apmi_scraper.py --inspect      # print the report forms' fields
Options:   --categories Equity Debt              # limit categories
           --services Discretionary              # limit service types
"""

import argparse
import json
import re
import sys
import time
from datetime import datetime, timezone
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup

BASE = "https://www.apmiindia.org/apmi/"
PERF_URL = BASE + "welcomeiaperformance.htm?action=PMSmenu"
TURN_URL = BASE + "IATurnoverReportUtility.htm?action=loadIATurnoverPage"

CATEGORIES = ["Equity", "Debt", "Hybrid", "Multi Asset"]
SERVICES = ["Discretionary", "Non Discretionary"]

HEADERS = {
    "User-Agent": "Mozilla/5.0 (PMS Tracker research script)",
    "Accept-Language": "en-IN,en;q=0.9",
}
DELAY_SECONDS = 3  # be polite between requests


# ---------------------------------------------------------------- helpers
def norm(text):
    return re.sub(r"\s+", " ", str(text or "")).strip().lower()


def clean_number(value):
    """'₹1,983.43' -> 1983.43, 'NA' -> None"""
    text = str(value or "").replace("₹", "").replace(",", "").strip()
    if text.upper() in {"", "NA", "N/A", "-", "--"}:
        return None
    try:
        return float(text)
    except ValueError:
        return text


def find_table(html, required, forbidden=()):
    soup = BeautifulSoup(html, "lxml")
    for table in soup.find_all("table"):
        rows = table.find_all("tr")
        if not rows:
            continue
        header_row = next((r for r in rows if r.find("th")), rows[0])
        headers = [norm(c.get_text()) for c in header_row.find_all(["th", "td"])]
        if all(any(r in h for h in headers) for r in required) and not any(
            any(b in h for h in headers) for b in forbidden
        ):
            data = []
            for r in rows[rows.index(header_row) + 1:]:
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
    "pms provider name": "Provider", "ia name": "Investment Approach",
    "aum (in inr cr.)": "AUM (INR Cr)", "1 month": "Return 1M (%)",
    "3 months": "Return 3M (%)", "6 months": "Return 6M (%)",
    "1 year": "Return 1Y (%)", "2 years": "Return 2Y (%)",
    "3 years": "Return 3Y (%)", "4 years": "Return 4Y (%)",
    "5 years": "Return 5Y (%)", "since inception": "Return Since Inception (%)",
}
TURN_COLS = {
    "pms provider name": "Provider", "ia name": "Investment Approach",
    "1 month": "Turnover 1M", "1 year": "Turnover 1Y",
}


def parse_report(html, required, forbidden, colmap):
    headers, rows = find_table(html, required, forbidden)
    if headers is None:
        return None  # no table on this page (e.g. no strategies in this category)
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


# ---------------------------------------------------------------- form handling
def find_report_form(soup):
    """The form that contains the category select (the one with an 'Equity' option)."""
    for form in soup.find_all("form"):
        for sel in form.find_all("select"):
            if any(norm(o.get_text()) == "equity" for o in sel.find_all("option")):
                return form
    return None


def option_matching(select, label):
    target = norm(label).replace(" ", "")
    for o in select.find_all("option"):
        if norm(o.get_text()).replace(" ", "").replace("-", "") == target.replace("-", ""):
            return o
    return None


def select_role(select):
    texts = {norm(o.get_text()) for o in select.find_all("option")}
    if "equity" in texts:
        return "category"
    if "discretionary" in texts:
        return "service"
    if any("tri" in t for t in texts):
        return "benchmark"
    if len(texts) > 20:
        return "filter"  # provider / investment approach lists
    return "other"


def month_variants(as_on):
    """'31/08/2026' -> candidate formats for the month field."""
    d = datetime.strptime(as_on, "%d/%m/%Y")
    return [d.strftime(f) for f in ("%m-%Y", "%m/%Y", "%Y-%m", "%b-%Y", "%B-%Y", "%b %Y", "%d/%m/%Y", "%d-%m-%Y")]


def build_payload(form, category, service, month_value):
    payload = {}
    month_field = None
    for el in form.find_all(["input", "select"]):
        name = el.get("name")
        if not name:
            continue
        if el.name == "select":
            role = select_role(el)
            if role == "category":
                opt = option_matching(el, category)
                if opt is None:
                    raise RuntimeError(f"Category '{category}' not in form")
                payload[name] = opt.get("value", opt.get_text(strip=True))
            elif role == "service":
                opt = option_matching(el, service)
                if opt is None:
                    raise RuntimeError(f"Service type '{service}' not in form")
                payload[name] = opt.get("value", opt.get_text(strip=True))
            elif role == "filter":
                # leave provider / approach unfiltered
                blank = next((o for o in el.find_all("option")
                              if o.get("value", "x") in ("", "0", "-1") or re.search(r"select|all", norm(o.get_text()))), None)
                if blank is not None:
                    payload[name] = blank.get("value", "")
            else:
                chosen = el.find("option", selected=True) or el.find("option")
                if chosen is not None:
                    payload[name] = chosen.get("value", chosen.get_text(strip=True))
        else:
            itype = (el.get("type") or "text").lower()
            if itype in ("submit", "button", "reset", "image"):
                continue
            if itype in ("checkbox", "radio"):
                if el.has_attr("checked"):
                    payload[name] = el.get("value", "on")
                continue
            if re.search(r"month|date|year|ason", name + " " + (el.get("id") or ""), re.I):
                month_field = name
                payload[name] = month_value
            else:
                payload[name] = el.get("value", "")
    # the site has a Submit button; include it in case the server checks for it
    btn = form.find(["input", "button"], attrs={"type": "submit"})
    if btn is not None and btn.get("name"):
        payload[btn["name"]] = btn.get("value", "Submit")
    return payload, month_field


def selected_label(html, role):
    """Which option the server shows as selected in the response (to verify the filter applied)."""
    soup = BeautifulSoup(html, "lxml")
    form = find_report_form(soup)
    if form is None:
        return None
    for sel in form.find_all("select"):
        if select_role(sel) == role:
            opt = sel.find("option", selected=True)
            return norm(opt.get_text()) if opt else None
    return None


def submit_report(session, page_url, default_html, category, service, as_on):
    soup = BeautifulSoup(default_html, "lxml")
    form = find_report_form(soup)
    if form is None:
        raise RuntimeError("Report form not found; run --inspect")
    action = urljoin(page_url, form.get("action") or page_url)
    method = (form.get("method") or "post").lower()

    for mv in month_variants(as_on):
        payload, month_field = build_payload(form, category, service, mv)
        time.sleep(DELAY_SECONDS)
        resp = session.request(method, action, timeout=60,
                               data=payload if method == "post" else None,
                               params=payload if method == "get" else None)
        resp.raise_for_status()
        html = resp.text
        got_as_on = extract_as_on(html)
        if got_as_on == as_on or month_field is None:
            cat = selected_label(html, "category")
            if cat is not None and cat.replace(" ", "") != norm(category).replace(" ", ""):
                raise RuntimeError(f"Server returned category '{cat}' instead of '{category}'")
            return html
    raise RuntimeError("Could not submit the month in any known format; run --inspect")


# ---------------------------------------------------------------- merging
def merge(perf, turn):
    perf = perf.copy() if perf is not None else pd.DataFrame(columns=["Provider", "Investment Approach"])
    turn = turn.copy() if turn is not None else pd.DataFrame(columns=list(TURN_COLS.values()))
    for df in (perf, turn):
        df["_key"] = df["Provider"].map(norm) + "|" + df["Investment Approach"].map(norm)
        df["_dup"] = df.groupby("_key").cumcount()
    combined = perf.merge(turn.drop(columns=["Provider", "Investment Approach"]),
                          on=["_key", "_dup"], how="outer", indicator=True)
    missing = combined["Provider"].isna()
    if missing.any():
        names = turn.set_index(["_key", "_dup"])[["Provider", "Investment Approach"]]
        idx = list(zip(combined.loc[missing, "_key"], combined.loc[missing, "_dup"]))
        combined.loc[missing, ["Provider", "Investment Approach"]] = names.loc[idx].values
    combined["Matched In"] = combined["_merge"].map(
        {"both": "Both reports", "left_only": "Performance only", "right_only": "Turnover only"})
    return combined.drop(columns=["_key", "_dup", "_merge"])


# ---------------------------------------------------------------- main
def fetch(session, url):
    resp = session.get(url, timeout=60)
    resp.raise_for_status()
    return resp.text


def describe_forms(html, label):
    soup = BeautifulSoup(html, "lxml")
    print(f"\n=== Forms on {label} ===")
    for i, form in enumerate(soup.find_all("form")):
        print(f"\nForm #{i}: action={form.get('action')!r} method={form.get('method')!r} onsubmit={form.get('onsubmit')!r}")
        for el in form.find_all(["input", "select", "button"]):
            if el.name == "select":
                opts = [(o.get("value"), o.get_text(strip=True)) for o in el.find_all("option")][:5]
                print(f"  select name={el.get('name')!r} id={el.get('id')!r} role={select_role(el)} options={opts}")
            else:
                print(f"  {el.name} type={el.get('type')!r} name={el.get('name')!r} id={el.get('id')!r} value={el.get('value')!r}")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--categories", nargs="*", default=CATEGORIES)
    ap.add_argument("--services", nargs="*", default=SERVICES)
    ap.add_argument("--inspect", action="store_true")
    ap.add_argument("--out", default="apmi_pms_data.xlsx")
    ap.add_argument("--json", help="also write JSON for the website, e.g. docs/data.json")
    args = ap.parse_args()

    session = requests.Session()
    session.headers.update(HEADERS)
    perf_default = fetch(session, PERF_URL)
    time.sleep(DELAY_SECONDS)
    turn_default = fetch(session, TURN_URL)

    if args.inspect:
        describe_forms(perf_default, "IA Performance")
        describe_forms(turn_default, "IA Turnover")
        return

    as_on = extract_as_on(perf_default)
    print(f"Latest month on APMI: {as_on}")
    frames, failures = [], []

    for category in args.categories:
        for service in args.services:
            label = f"{category} / {service}"
            try:
                perf_html = submit_report(session, PERF_URL, perf_default, category, service, as_on)
                turn_html = submit_report(session, TURN_URL, turn_default, category, service, as_on)
                perf, turn = parse_performance(perf_html), parse_turnover(turn_html)
                if perf is None and turn is None:
                    print(f"  {label}: no strategies listed")
                    continue
                combined = merge(perf, turn)
                combined.insert(0, "Service Type", service)
                combined.insert(0, "Category", category)
                frames.append(combined)
                print(f"  {label}: {len(combined)} strategies")
            except Exception as e:
                failures.append(label)
                print(f"  {label}: FAILED ({e})", file=sys.stderr)

    # Fallback: if every form submission failed, keep the default page (Equity) so the site still updates
    if not frames:
        print("All form submissions failed; saving the default report only. Run --inspect and share the output.")
        combined = merge(parse_performance(perf_default), parse_turnover(turn_default))
        combined.insert(0, "Service Type", "Default view")
        combined.insert(0, "Category", "Equity")
        frames.append(combined)

    data = pd.concat(frames, ignore_index=True)
    data.insert(0, "As On", as_on)

    with pd.ExcelWriter(args.out, engine="openpyxl") as xl:
        data.to_excel(xl, sheet_name="All strategies", index=False)
        ws = xl.sheets["All strategies"]
        ws.freeze_panes = "F2"
        ws.auto_filter.ref = ws.dimensions
    print(f"Saved {args.out} ({len(data)} rows)")

    if args.json:
        clean = data.astype(object).where(pd.notna(data), None)
        payload = {
            "as_on": as_on,
            "updated_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC"),
            "source": "Association of Portfolio Managers in India (apmiindia.org)",
            "failed": failures,
            "rows": clean.to_dict(orient="records"),
        }
        with open(args.json, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False)
        print(f"Saved {args.json}")

    if failures:
        print(f"\nThese combinations failed: {', '.join(failures)}")


if __name__ == "__main__":
    main()
