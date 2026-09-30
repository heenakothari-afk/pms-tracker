"""
APMI PMS scraper: returns, AUM and portfolio turnover (1M / 1Y)
for every category (Equity, Debt, Hybrid, Multi Asset), service type
(Discretionary, Non Discretionary) and month-end.

Source: https://www.apmiindia.org (IA Performance Report + IA Turnover Report)

Install:  pip install requests beautifulsoup4 pandas openpyxl lxml
Run:      python apmi_scraper.py --site docs            # latest month + backfill history
          python apmi_scraper.py --site docs --backfill 0   # latest month only
          python apmi_scraper.py --inspect               # print what the report pages contain

Writes (inside --site folder):
  data.json               latest month (what the website opens by default)
  history/YYYY-MM.json    one file per month-end
  months.json             list of available months
  apmi_pms_data.xlsx      latest month as a spreadsheet
"""

import argparse
import calendar
import json
import os
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
HEADERS = {"User-Agent": "Mozilla/5.0 (PMS Tracker research script)", "Accept-Language": "en-IN,en;q=0.9"}
DELAY_SECONDS = 3
_diagnosed = set()


# ================================================================ small helpers
def norm(text):
    return re.sub(r"\s+", " ", str(text or "")).strip().lower()


def squash(text):
    return re.sub(r"[^a-z0-9]", "", norm(text))


def clean_number(value):
    text = str(value or "").replace("₹", "").replace(",", "").strip()
    if text.upper() in {"", "NA", "N/A", "-", "--"}:
        return None
    try:
        return float(text)
    except ValueError:
        return text


def extract_as_on(html):
    m = re.search(r"As on\s*(\d{2}/\d{2}/\d{4})", BeautifulSoup(html, "lxml").get_text(" "))
    return m.group(1) if m else None


def month_id(as_on):                      # '31/08/2026' -> '2026-08'
    d = datetime.strptime(as_on, "%d/%m/%Y")
    return d.strftime("%Y-%m")


def month_end(mid):                       # '2026-08' -> '31/08/2026'
    y, m = map(int, mid.split("-"))
    return f"{calendar.monthrange(y, m)[1]:02d}/{m:02d}/{y}"


def prev_month(mid):
    y, m = map(int, mid.split("-"))
    return f"{y - 1}-12" if m == 1 else f"{y}-{m - 1:02d}"


def month_variants(mid):
    d = datetime.strptime(mid + "-01", "%Y-%m-%d")
    last = month_end(mid)
    out = [d.strftime(f) for f in ("%m-%Y", "%m/%Y", "%Y-%m", "%b-%Y", "%B-%Y", "%b %Y", "%B %Y", "%b-%y")]
    out += [last, last.replace("/", "-")]
    return out


# ================================================================ tables
PERF_COLS = {
    "pms provider name": "Provider", "ia name": "Investment Approach",
    "aum (in inr cr.)": "AUM (INR Cr)", "1 month": "Return 1M (%)",
    "3 months": "Return 3M (%)", "6 months": "Return 6M (%)",
    "1 year": "Return 1Y (%)", "2 years": "Return 2Y (%)",
    "3 years": "Return 3Y (%)", "4 years": "Return 4Y (%)",
    "5 years": "Return 5Y (%)", "since inception": "Return Since Inception (%)",
}
TURN_COLS = {"pms provider name": "Provider", "ia name": "Investment Approach",
             "1 month": "Turnover 1M", "1 year": "Turnover 1Y"}


def find_table(html, required, forbidden=()):
    soup = BeautifulSoup(html, "lxml")
    for table in soup.find_all("table"):
        rows = table.find_all("tr")
        if not rows:
            continue
        header_row = next((r for r in rows if r.find("th")), rows[0])
        headers = [norm(c.get_text()) for c in header_row.find_all(["th", "td"])]
        if all(any(r in h for h in headers) for r in required) and not any(any(b in h for h in headers) for b in forbidden):
            data = [[c.get_text(" ", strip=True) for c in r.find_all("td")] for r in rows[rows.index(header_row) + 1:]]
            return headers, [d for d in data if len(d) == len(headers)]
    return None, None


def parse_report(html, required, forbidden, colmap):
    headers, rows = find_table(html, required, forbidden)
    if headers is None:
        return None
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


def ia_names(df):
    return set() if df is None else set(df["Investment Approach"].map(norm))


# ================================================================ report form (dropdowns, radios, or loose fields)
def label_for(el, soup):
    """Visible text next to a radio/checkbox."""
    if el.get("id"):
        lab = soup.find("label", attrs={"for": el["id"]})
        if lab:
            return lab.get_text(" ", strip=True)
    if el.parent is not None and el.parent.name == "label":
        return el.parent.get_text(" ", strip=True)
    sib = el.next_sibling
    while sib is not None and getattr(sib, "name", None) not in ("input", "select", "br"):
        text = sib.get_text(" ", strip=True) if hasattr(sib, "get_text") else str(sib).strip()
        if text:
            return text
        sib = sib.next_sibling
    return el.get("value", "")


def role_of(labels):
    texts = {squash(t) for t in labels}
    if "equity" in texts or "debt" in texts:
        return "category"
    if "discretionary" in texts or "nondiscretionary" in texts:
        return "service"
    if any(re.search(r"(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\d{2,4}|\d{1,2}\d{4}", t) for t in texts) and len(texts) > 3:
        return "month"
    if any("tri" in norm(t).split() or "nifty" in norm(t) or "bse" in norm(t) or "crisil" in norm(t) for t in labels):
        return "benchmark"
    if len(texts) > 20:
        return "filter"
    return "other"


class ReportForm:
    def __init__(self, html, page_url):
        soup = BeautifulSoup(html, "lxml")
        self.soup = soup
        container = None
        for form in soup.find_all("form"):
            if re.search(r"\bequity\b", norm(form.get_text(" "))) or any(
                    squash(label_for(i, soup)) == "equity" for i in form.find_all("input")):
                container = form
                break
        self.is_form = container is not None
        self.container = container or soup.body or soup
        self.action = urljoin(page_url, (container.get("action") if container else None) or page_url)
        self.method = ((container.get("method") if container else None) or "post").lower()
        self.fields = []   # (kind, name, role, options[(value,label)], default)
        seen_radio = {}
        for el in self.container.find_all(["select", "input", "textarea"]):
            name = el.get("name")
            if not name:
                continue
            if el.name == "select":
                opts = [(o.get("value", o.get_text(strip=True)), o.get_text(" ", strip=True)) for o in el.find_all("option")]
                sel = el.find("option", selected=True)
                default = (sel.get("value", sel.get_text(strip=True)) if sel else (opts[0][0] if opts else ""))
                self.fields.append(["select", name, role_of([o[1] for o in opts] + [o[0] for o in opts]), opts, default])
                continue
            itype = (el.get("type") or "text").lower()
            if itype in ("submit", "button", "reset", "image", "file"):
                continue
            if itype in ("radio", "checkbox"):
                lab = label_for(el, soup)
                if name not in seen_radio:
                    seen_radio[name] = ["radio" if itype == "radio" else "checkbox", name, None, [], None]
                    self.fields.append(seen_radio[name])
                f = seen_radio[name]
                f[3].append((el.get("value", "on"), lab))
                if el.has_attr("checked"):
                    f[4] = el.get("value", "on")
                continue
            ident = f"{name} {el.get('id', '')} {el.get('placeholder', '')}"
            role = "month" if re.search(r"month|date|year|ason|period", ident, re.I) else "other"
            self.fields.append(["input", name, role, [], el.get("value", "")])
        for f in self.fields:
            if f[0] in ("radio", "checkbox"):
                f[2] = role_of([o[1] for o in f[3]] + [o[0] for o in f[3]])
        self.month_format = None

    def has(self, role):
        return any(f[2] == role for f in self.fields)

    def month_options(self):
        """Months offered by a month dropdown, as 'YYYY-MM' ids (empty if the month is a text box)."""
        out = []
        for f in self.fields:
            if f[0] == "select" and f[2] == "month":
                for value, label in f[3]:
                    for text in (label, value):
                        for fmt in ("%b-%Y", "%B-%Y", "%b %Y", "%B %Y", "%m-%Y", "%m/%Y", "%Y-%m", "%d/%m/%Y", "%d-%m-%Y", "%b-%y"):
                            try:
                                out.append(datetime.strptime(text.strip(), fmt).strftime("%Y-%m"))
                                break
                            except ValueError:
                                continue
                        else:
                            continue
                        break
        return sorted(set(out), reverse=True)

    @staticmethod
    def _pick(options, wanted):
        w = squash(wanted)
        for value, label in options:
            if squash(label) == w or squash(value) == w:
                return value
        return None

    def payloads(self, category, service, mid):
        """Yield candidate payloads (one per month format when the month is a text box)."""
        variants = month_variants(mid)
        order = [self.month_format] if self.month_format is not None else range(len(variants))
        for mfmt in order:
            data = {}
            for kind, name, role, opts, default in self.fields:
                if role == "category":
                    v = self._pick(opts, category)
                    if v is None:
                        raise RuntimeError(f"'{category}' is not an option on APMI's form")
                    data[name] = v
                elif role == "service":
                    v = self._pick(opts, service)
                    if v is None:
                        raise RuntimeError(f"'{service}' is not an option on APMI's form")
                    data[name] = v
                elif role == "month":
                    if kind == "select":
                        v = next((val for val, lab in opts for var in month_variants(mid)
                                  if squash(lab) == squash(var) or squash(val) == squash(var)), None)
                        if v is None:
                            raise RuntimeError(f"month {mid} not offered by APMI")
                        data[name] = v
                    else:
                        data[name] = variants[mfmt]
                elif role == "filter" and kind == "select":
                    blank = next((val for val, lab in opts if val in ("", "0", "-1") or re.search(r"select|all", norm(lab))), None)
                    if blank is not None:
                        data[name] = blank
                elif kind in ("radio", "checkbox"):
                    if default is not None:
                        data[name] = default
                elif default is not None:
                    data[name] = default
            yield mfmt, data
            if not any(f[2] == "month" and f[0] == "input" for f in self.fields):
                break  # month isn't a text box, so one attempt is enough


def diagnose(html, label):
    """Print what's on the page so the form can be matched precisely."""
    if label in _diagnosed:
        return
    _diagnosed.add(label)
    soup = BeautifulSoup(html, "lxml")
    print(f"\n----- DIAGNOSTICS: {label} -----")
    print(f"forms on page: {len(soup.find_all('form'))}")
    for i, form in enumerate(soup.find_all("form")):
        print(f"form #{i}: action={form.get('action')!r} method={form.get('method')!r} "
              f"onsubmit={str(form.get('onsubmit'))[:80]!r} id={form.get('id')!r} name={form.get('name')!r}")
    for el in soup.find_all(["select", "input", "button"])[:60]:
        if el.name == "select":
            opts = [(o.get("value"), o.get_text(strip=True)) for o in el.find_all("option")][:6]
            print(f"  select name={el.get('name')!r} id={el.get('id')!r} onchange={str(el.get('onchange'))[:60]!r} options={opts}")
        else:
            print(f"  {el.name} type={el.get('type')!r} name={el.get('name')!r} id={el.get('id')!r} "
                  f"value={str(el.get('value'))[:40]!r} onclick={str(el.get('onclick'))[:60]!r} label={label_for(el, soup)[:30]!r}")
    text = str(soup)
    i = text.lower().find("equity")
    if i >= 0:
        print("  html around 'Equity':", re.sub(r"\s+", " ", text[max(0, i - 400): i + 400]))
    for s in soup.find_all("script"):
        body = s.string or ""
        if re.search(r"submit|\.htm|ajax|action", body, re.I):
            print("  script:", re.sub(r"\s+", " ", body)[:500])
    print("----- END DIAGNOSTICS -----\n")


# ================================================================ fetching one report
def fetch(session, url):
    resp = session.get(url, timeout=60)
    resp.raise_for_status()
    return resp.text


def submit(session, form, category, service, mid):
    want = month_end(mid)
    last_as_on = None
    for mfmt, data in form.payloads(category, service, mid):
        time.sleep(DELAY_SECONDS)
        if form.method == "get":
            resp = session.get(form.action, params=data, timeout=60)
        else:
            resp = session.post(form.action, data=data, timeout=60)
        resp.raise_for_status()
        last_as_on = extract_as_on(resp.text)
        if last_as_on == want:
            form.month_format = mfmt
            return resp.text
    raise RuntimeError(f"APMI returned data as on {last_as_on} instead of {want}")


def scrape_month(session, perf_form, turn_form, mid, default_perf, is_latest):
    frames, failures = [], []
    default_equity = ia_names(default_perf)
    for category in CATEGORIES:
        for service in SERVICES:
            label = f"{mid} {category} / {service}"
            try:
                perf_html = submit(session, perf_form, category, service, mid)
                turn_html = submit(session, turn_form, category, service, mid)
                perf, turn = parse_performance(perf_html), parse_turnover(turn_html)
                if perf is None and turn is None:
                    print(f"  {label}: no strategies listed")
                    continue
                # Guard: if a non-equity request returns exactly the equity list, the filter was ignored
                if is_latest and category != "Equity" and perf is not None and default_equity and ia_names(perf) == default_equity:
                    raise RuntimeError("APMI ignored the category and returned the Equity list")
                combined = merge(perf, turn)
                combined.insert(0, "Service Type", service)
                combined.insert(0, "Category", category)
                frames.append(combined)
                print(f"  {label}: {len(combined)} strategies")
            except Exception as e:
                failures.append(f"{category} / {service}")
                print(f"  {label}: FAILED ({e})")
    return frames, failures


# ================================================================ merge + save
def merge(perf, turn):
    perf = perf.copy() if perf is not None else pd.DataFrame(columns=["Provider", "Investment Approach"])
    turn = turn.copy() if turn is not None else pd.DataFrame(columns=list(TURN_COLS.values()))
    for df in (perf, turn):
        df["_key"] = df["Provider"].map(norm) + "|" + df["Investment Approach"].map(norm)
        df["_dup"] = df.groupby("_key").cumcount()
    combined = perf.merge(turn.drop(columns=["Provider", "Investment Approach"]), on=["_key", "_dup"], how="outer", indicator=True)
    missing = combined["Provider"].isna()
    if missing.any():
        names = turn.set_index(["_key", "_dup"])[["Provider", "Investment Approach"]]
        idx = list(zip(combined.loc[missing, "_key"], combined.loc[missing, "_dup"]))
        combined.loc[missing, ["Provider", "Investment Approach"]] = names.loc[idx].values
    combined["Matched In"] = combined["_merge"].map({"both": "Both reports", "left_only": "Performance only", "right_only": "Turnover only"})
    return combined.drop(columns=["_key", "_dup", "_merge"])


def month_payload(data, as_on, failures):
    clean = data.astype(object).where(pd.notna(data), None)
    return {
        "as_on": as_on,
        "updated_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC"),
        "source": "Association of Portfolio Managers in India (apmiindia.org)",
        "failed": failures,
        "rows": clean.to_dict(orient="records"),
    }


def write_json(path, obj):
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False)


def rebuild_index(site):
    hist = os.path.join(site, "history")
    months = []
    for fn in sorted(os.listdir(hist), reverse=True) if os.path.isdir(hist) else []:
        if re.fullmatch(r"\d{4}-\d{2}\.json", fn):
            with open(os.path.join(hist, fn), encoding="utf-8") as f:
                d = json.load(f)
            cats = sorted({r.get("Category") or "Equity" for r in d.get("rows", [])})
            months.append({"id": fn[:-5], "as_on": d.get("as_on"), "strategies": len(d.get("rows", [])), "categories": cats})
    write_json(os.path.join(site, "months.json"), {"latest": months[0]["id"] if months else None, "months": months})
    return months


def history_is_complete(site, mid):
    path = os.path.join(site, "history", f"{mid}.json")
    if not os.path.exists(path):
        return False
    with open(path, encoding="utf-8") as f:
        d = json.load(f)
    return not d.get("failed")


# ================================================================ main
def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--site", default="docs", help="website folder to write into")
    ap.add_argument("--backfill", type=int, default=24, help="max older months to add this run (0 = none)")
    ap.add_argument("--earliest", default="2019-01", help="don't go back before this month (YYYY-MM)")
    ap.add_argument("--inspect", action="store_true")
    args = ap.parse_args()

    session = requests.Session()
    session.headers.update(HEADERS)
    perf_default = fetch(session, PERF_URL)
    time.sleep(DELAY_SECONDS)
    turn_default = fetch(session, TURN_URL)

    if args.inspect:
        diagnose(perf_default, "IA Performance page")
        diagnose(turn_default, "IA Turnover page")
        return

    latest_as_on = extract_as_on(perf_default)
    latest = month_id(latest_as_on)
    print(f"Latest month on APMI: {latest_as_on}")

    perf_form, turn_form = ReportForm(perf_default, PERF_URL), ReportForm(turn_default, TURN_URL)
    print(f"Performance page: form={'yes' if perf_form.is_form else 'no (loose fields)'}; "
          f"category={perf_form.has('category')} service={perf_form.has('service')} month={perf_form.has('month')}")
    print(f"Turnover page:    form={'yes' if turn_form.is_form else 'no (loose fields)'}; "
          f"category={turn_form.has('category')} service={turn_form.has('service')} month={turn_form.has('month')}")
    forms_ok = perf_form.has("category") and turn_form.has("category")
    if not forms_ok:
        diagnose(perf_default, "IA Performance page")
        diagnose(turn_default, "IA Turnover page")

    default_perf = parse_performance(perf_default)

    # ---------- latest month
    frames, failures = ([], CATEGORIES)
    if forms_ok:
        print(f"\nLatest month {latest}:")
        frames, failures = scrape_month(session, perf_form, turn_form, latest, default_perf, True)
        if failures:
            diagnose(perf_default, "IA Performance page")
    if not frames:
        print("Category requests failed; saving APMI's default view so the site still updates.")
        data = merge(default_perf, parse_turnover(turn_default))
        data.insert(0, "Service Type", "Default view")
        data.insert(0, "Category", "Equity")
        failures = [f"{c} / {s}" for c in CATEGORIES for s in SERVICES]
    else:
        data = pd.concat(frames, ignore_index=True)
    data.insert(0, "As On", latest_as_on)

    payload = month_payload(data, latest_as_on, failures)
    write_json(os.path.join(args.site, "data.json"), payload)
    write_json(os.path.join(args.site, "history", f"{latest}.json"), payload)
    with pd.ExcelWriter(os.path.join(args.site, "apmi_pms_data.xlsx"), engine="openpyxl") as xl:
        data.to_excel(xl, sheet_name="Latest month", index=False)
        xl.sheets["Latest month"].freeze_panes = "F2"
        xl.sheets["Latest month"].auto_filter.ref = xl.sheets["Latest month"].dimensions
    print(f"Saved latest month: {len(data)} strategies")

    # ---------- older months
    if forms_ok and frames and args.backfill > 0 and perf_form.has("month"):
        offered = perf_form.month_options()
        candidates = offered[1:] if offered else []
        if not candidates:
            m = prev_month(latest)
            while m >= args.earliest:
                candidates.append(m)
                m = prev_month(m)
        done, empty_streak = 0, 0
        for mid in candidates:
            if mid < args.earliest or done >= args.backfill:
                break
            if history_is_complete(args.site, mid):
                continue
            print(f"\nHistory {mid}:")
            try:
                f_hist, fail_hist = scrape_month(session, perf_form, turn_form, mid, default_perf, False)
            except Exception as e:
                print(f"  {mid}: FAILED ({e})")
                f_hist, fail_hist = [], ["all"]
            if not f_hist:
                empty_streak += 1
                if empty_streak >= 3 and not offered:
                    print("  Three months in a row with no data; assuming APMI's history ends here.")
                    break
                continue
            empty_streak = 0
            hist = pd.concat(f_hist, ignore_index=True)
            hist.insert(0, "As On", month_end(mid))
            write_json(os.path.join(args.site, "history", f"{mid}.json"), month_payload(hist, month_end(mid), fail_hist))
            done += 1
            print(f"  saved {mid}: {len(hist)} strategies")
    elif args.backfill > 0:
        print("\nSkipping history backfill until category/month requests work.")

    months = rebuild_index(args.site)
    print(f"\nMonths available on the site: {', '.join(m['id'] for m in months)}")
    if failures:
        print(f"Failed this run: {', '.join(failures)}")


if __name__ == "__main__":
    main()
