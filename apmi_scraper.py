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
from urllib.parse import urljoin, urlparse

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
    """APMI's report form: radio buttons for category/service, month+year dropdowns filled by JavaScript,
    and a hidden asOnDate that the Submit button fills in."""

    def __init__(self, html, page_url, session=None):
        soup = BeautifulSoup(html, "lxml")
        self.soup, self.page_url, self.session = soup, page_url, session
        container = None
        for form in soup.find_all("form"):
            if form.find("input", attrs={"type": "radio"}) and re.search(r"\bequity\b", norm(form.get_text(" "))):
                container = form
                break
        self.is_form = container is not None
        self.container = container or soup.body or soup
        self.form_action = urljoin(page_url, (container.get("action") if container else "") or page_url)
        self.method = ((container.get("method") if container else None) or "post").lower()
        self.fields = []   # [kind, name, role, options[(value,label)], default]
        radios = {}
        for el in self.container.find_all(["select", "input", "textarea"]):
            name = el.get("name")
            if not name:
                continue
            if el.name == "select":
                opts = [(o.get("value", o.get_text(strip=True)), o.get_text(" ", strip=True)) for o in el.find_all("option")]
                sel = el.find("option", selected=True)
                default = sel.get("value", sel.get_text(strip=True)) if sel else None
                role = role_of([o[1] for o in opts] + [o[0] for o in opts]) if opts else "other"
                if re.search(r"month", name, re.I) and not re.search(r"year", name, re.I):
                    role = "month_part"
                elif re.search(r"year", name, re.I):
                    role = "year_part"
                self.fields.append(["select", name, role, opts, default])
                continue
            itype = (el.get("type") or "text").lower()
            if itype in ("submit", "button", "reset", "image", "file"):
                continue
            if itype in ("radio", "checkbox"):
                if name not in radios:
                    radios[name] = [itype, name, None, [], None]
                    self.fields.append(radios[name])
                radios[name][3].append((el.get("value", "on"), label_for(el, soup)))
                if el.has_attr("checked"):
                    radios[name][4] = el.get("value", "on")
                continue
            ident = f"{name} {el.get('id', '')}"
            role = "date" if re.search(r"date|ason", ident, re.I) else "other"
            self.fields.append(["input", name, role, [], el.get("value", "")])
        for f in self.fields:
            if f[0] in ("radio", "checkbox"):
                f[2] = role_of([o[1] for o in f[3]] + [o[0] for o in f[3]])
        self.js = self._collect_js(html)
        self.js_action = self._action_from_js()
        self.locked = None   # (action, month_variant_index) once a request works

    # ----- JavaScript: find the submit function to learn the real target URL
    def _collect_js(self, html):
        chunks = [s.string or "" for s in self.soup.find_all("script") if not s.get("src")]
        if self.session is not None:
            for s in self.soup.find_all("script", src=True):
                src = urljoin(self.page_url, s["src"])
                if urlparse(src).netloc != urlparse(self.page_url).netloc or re.search(r"jquery|bootstrap|popper|datatable|select2|moment|chart", src, re.I):
                    continue
                try:
                    r = self.session.get(src, timeout=30)
                    if r.ok:
                        chunks.append(r.text)
                except requests.RequestException:
                    pass
        return "\n".join(chunks)

    def js_function(self, name):
        m = re.search(r"function\s+" + re.escape(name) + r"\s*\([^)]*\)\s*\{", self.js)
        if not m:
            return None
        depth, i = 0, m.end() - 1
        while i < len(self.js):
            if self.js[i] == "{":
                depth += 1
            elif self.js[i] == "}":
                depth -= 1
                if depth == 0:
                    return self.js[m.start(): i + 1]
            i += 1
        return self.js[m.start(): m.start() + 3000]

    def _action_from_js(self):
        body = self.js_function("onClickSubmit") or ""
        m = re.search(r"""\.action\s*=\s*["']([^"']+)["']""", body) or re.search(r"""(?:url|href)\s*[:=]\s*["']([^"']+\.htm[^"']*)["']""", body)
        return urljoin(self.page_url, m.group(1)) if m else None

    def has(self, role):
        return any(f[2] == role for f in self.fields)

    def month_ready(self):
        return self.has("date") or self.has("month_part") or self.has("month")

    @staticmethod
    def _pick(options, wanted):
        w = squash(wanted)
        for value, label in options:
            if squash(label) == w or squash(value) == w:
                return value
        return None

    @staticmethod
    def month_variants(mid):
        y, m = mid.split("-")
        d = datetime(int(y), int(m), 1)
        last = month_end(mid)
        mm_parts = [m, str(int(m)), d.strftime("%b"), d.strftime("%B")]
        dates = [last, f"{m}/{y}", last.replace("/", "-"), f"{y}-{m}-{last[:2]}", f"{m}-{y}", d.strftime("%b-%Y")]
        # most likely first: two-digit month + dd/mm/yyyy (the format APMI prints on the page)
        return [(mp, y, dt) for dt in dates for mp in mm_parts]

    def actions(self):
        base = self.page_url.split("?")[0]
        out = [a for a in (self.js_action, self.form_action, self.page_url, base) if a]
        return list(dict.fromkeys(out))

    def payload(self, category, service, mid, variant):
        month_part, year_part, date_value = self.month_variants(mid)[variant]
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
            elif role == "month_part":
                data[name] = month_part
            elif role == "year_part":
                data[name] = year_part
            elif role == "date":
                data[name] = date_value
            elif role == "month" and kind == "select":
                v = next((val for val, lab in opts for var in (date_value, month_end(mid)) if squash(lab) == squash(var) or squash(val) == squash(var)), None)
                if v is not None:
                    data[name] = v
            elif role == "filter":
                blank = next((val for val, lab in opts if val in ("", "0", "-1") or re.search(r"select|all", norm(lab))), None)
                if blank is not None:
                    data[name] = blank
            elif kind in ("radio", "checkbox"):
                if default is not None:
                    data[name] = default
            elif default is not None:
                data[name] = default
        return data

    def post(self, session, action, data):
        time.sleep(DELAY_SECONDS)
        if self.method == "get":
            r = session.get(action, params=data, timeout=60)
        else:
            r = session.post(action, data=data, timeout=60, headers={"Referer": self.page_url})
        r.raise_for_status()
        return r.text

    def probe(self, session, mid, parse, equity_names):
        """Find an (action, month format) that returns genuine Debt data for this month."""
        want = month_end(mid)
        n = len(self.month_variants(mid))
        for action in self.actions():
            for variant in range(n):
                try:
                    html = self.post(session, action, self.payload("Debt", "Discretionary", mid, variant))
                except Exception as e:
                    print(f"    probe {action} #{variant}: error {e}")
                    continue
                got = extract_as_on(html)
                df = parse(html)
                names = ia_names(df)
                ok = got == want and df is not None and len(df) and names != equity_names
                if ok:
                    self.locked = (action, variant)
                    print(f"    probe OK: {action} with month format #{variant} {self.month_variants(mid)[variant]}")
                    return True
            print(f"    probe: {action} didn't return Debt data")
        return False

    def fetch_report(self, session, category, service, mid):
        if self.locked is None:
            raise RuntimeError("form not working yet")
        action, variant = self.locked
        html = self.post(session, action, self.payload(category, service, mid, variant))
        got = extract_as_on(html)
        if got != month_end(mid):
            raise RuntimeError(f"APMI returned data as on {got} instead of {month_end(mid)}")
        return html


def fetch(session, url):
    resp = session.get(url, timeout=60)
    resp.raise_for_status()
    return resp.text


def diagnose(html, label, form=None):
    """Print what's on the page (and the submit JavaScript) so the form can be matched precisely."""
    if label in _diagnosed:
        return
    _diagnosed.add(label)
    soup = BeautifulSoup(html, "lxml")
    print(f"\n----- DIAGNOSTICS: {label} -----")
    for i, f in enumerate(soup.find_all("form")):
        print(f"form #{i}: action={f.get('action')!r} method={f.get('method')!r} name={f.get('name')!r}")
    for el in soup.find_all(["select", "input"])[:40]:
        if el.name == "select":
            opts = [(o.get("value"), o.get_text(strip=True)) for o in el.find_all("option")][:4]
            print(f"  select name={el.get('name')!r} onchange={str(el.get('onchange'))[:90]!r} options={opts}")
        else:
            print(f"  input type={el.get('type')!r} name={el.get('name')!r} value={str(el.get('value'))[:30]!r} onchange={str(el.get('onchange'))[:60]!r}")
    print("  scripts:", [s.get("src") for s in soup.find_all("script", src=True)])
    if form is not None:
        print(f"  form actions tried: {form.actions()}")
        for fn in ("onClickSubmit", "PerformanceMonth", "onChangeMonth", "strategytype", "getSelectedData"):
            body = form.js_function(fn)
            print(f"\n  JS {fn}:", re.sub(r"\s+", " ", body)[:2500] if body else "(not found)")
    print("----- END DIAGNOSTICS -----\n")


def scrape_month(session, perf_form, turn_form, mid):
    frames, failures, seen = [], [], {}
    for category in CATEGORIES:
        for service in SERVICES:
            label = f"{mid} {category} / {service}"
            try:
                perf = parse_performance(perf_form.fetch_report(session, category, service, mid))
                turn = parse_turnover(turn_form.fetch_report(session, category, service, mid))
                if perf is None and turn is None:
                    print(f"  {label}: no strategies listed")
                    continue
                names = frozenset(ia_names(perf if perf is not None else turn))
                if names and names in seen:
                    raise RuntimeError(f"same list as {seen[names]}; APMI ignored the filter")
                seen[names] = f"{category} / {service}"
                combined = merge(perf, turn)
                combined.insert(0, "Service Type", service)
                combined.insert(0, "Category", category)
                frames.append(combined)
                print(f"  {label}: {len(combined)} strategies")
            except Exception as e:
                failures.append(f"{category} / {service}")
                print(f"  {label}: FAILED ({e})")
                if not frames and category == "Equity" and service == SERVICES[0] and "as on None" in str(e):
                    print(f"  {mid}: APMI has no data for this month")
                    return [], ["no data"]
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
        diagnose(perf_default, "IA Performance page", ReportForm(perf_default, PERF_URL, session))
        diagnose(turn_default, "IA Turnover page", ReportForm(turn_default, TURN_URL, session))
        return

    latest_as_on = extract_as_on(perf_default)
    latest = month_id(latest_as_on)
    print(f"Latest month on APMI: {latest_as_on}")

    perf_form = ReportForm(perf_default, PERF_URL, session)
    turn_form = ReportForm(turn_default, TURN_URL, session)
    for label, f in (("Performance", perf_form), ("Turnover", turn_form)):
        print(f"{label} page: category={f.has('category')} service={f.has('service')} "
              f"month={f.month_ready()} submit-script-target={f.js_action}")

    default_perf = parse_performance(perf_default)
    equity_names = ia_names(default_perf)

    print("\nChecking how APMI's form accepts requests:")
    forms_ok = (perf_form.has("category") and turn_form.has("category")
                and perf_form.probe(session, latest, parse_performance, equity_names)
                and turn_form.probe(session, latest, parse_turnover, ia_names(parse_turnover(turn_default))))
    if not forms_ok:
        diagnose(perf_default, "IA Performance page", perf_form)
        diagnose(turn_default, "IA Turnover page", turn_form)

    # ---------- latest month
    frames, failures = [], []
    if forms_ok:
        print(f"\nLatest month {latest}:")
        frames, failures = scrape_month(session, perf_form, turn_form, latest)
    if not frames:
        print("Category requests aren't working yet; saving APMI's default Equity view so the site still updates.")
        data = merge(default_perf, parse_turnover(turn_default))
        data.insert(0, "Service Type", "Default view")
        data.insert(0, "Category", "Equity")
        failures = ["default-only"]
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
    if forms_ok and frames and args.backfill > 0:
        done, empty_streak, mid = 0, 0, prev_month(latest)
        while mid >= args.earliest and done < args.backfill:
            if not history_is_complete(args.site, mid):
                print(f"\nHistory {mid}:")
                f_hist, fail_hist = scrape_month(session, perf_form, turn_form, mid)
                if f_hist:
                    empty_streak = 0
                    hist = pd.concat(f_hist, ignore_index=True)
                    hist.insert(0, "As On", month_end(mid))
                    write_json(os.path.join(args.site, "history", f"{mid}.json"), month_payload(hist, month_end(mid), fail_hist))
                    done += 1
                    print(f"  saved {mid}: {len(hist)} strategies")
                else:
                    empty_streak += 1
                    if empty_streak >= 3:
                        print("  Three months in a row with no data; assuming APMI's history ends here.")
                        break
            mid = prev_month(mid)
    elif args.backfill > 0:
        print("\nSkipping history until category requests work.")

    months = rebuild_index(args.site)
    print(f"\nMonths available on the site: {', '.join(m['id'] for m in months)}")
    if failures:
        print(f"Not collected this run: {', '.join(failures)}")


if __name__ == "__main__":
    main()
