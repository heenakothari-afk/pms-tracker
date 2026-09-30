"""
APMI PMS scraper: returns, AUM and portfolio turnover (1M / 1Y)
for every category (Equity, Debt, Hybrid, Multi Asset), service type
(Discretionary, Non Discretionary) and month-end.

Source: https://www.apmiindia.org (IA Performance Report + IA Turnover Report)

Install:  pip install requests beautifulsoup4 pandas openpyxl lxml
Run:      python apmi_scraper.py --site docs            # latest month + backfill history
          python apmi_scraper.py --site docs --backfill 0   # latest month only

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
DELAY_SECONDS = 2
_diagnosed = set()


# ================================================================ small helpers
def norm(text):
    return re.sub(r"\s+", " ", str(text or "")).strip().lower()


def squash(text):
    return re.sub(r"[^a-z0-9]", "", norm(text))


def clean_number(value):
    text = str(value or "").replace(",", "").strip()
    text = re.sub(r"^[^\d\-.NAna/]+", "", text)   # drop ₹ (or a garbled version of it)
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


# ================================================================ APMI's background (AJAX) requests
AJAX_HEADERS = {"X-Requested-With": "XMLHttpRequest",
                "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"}


def js_sources(session, page_url, html):
    soup = BeautifulSoup(html, "lxml")
    chunks = [s.string or "" for s in soup.find_all("script") if not s.get("src")]
    for s in soup.find_all("script", src=True):
        src = urljoin(page_url, s["src"])
        if urlparse(src).netloc != urlparse(page_url).netloc:
            continue
        if re.search(r"jquery-|bootstrap|popper|datatables|moment|multiselect|bootbox|topbookmark", src, re.I):
            continue
        try:
            r = session.get(src, timeout=30)
            if r.ok:
                chunks.append(f"/* {src} */\n" + r.text)
        except requests.RequestException:
            pass
    return "\n".join(chunks)


def js_function(js, name):
    m = re.search(r"function\s+" + re.escape(name) + r"\s*\([^)]*\)\s*\{", js)
    if not m:
        return None
    depth, i = 0, m.end() - 1
    while i < len(js):
        if js[i] == "{":
            depth += 1
        elif js[i] == "}":
            depth -= 1
            if depth == 0:
                return js[m.start(): i + 1]
        i += 1
    return js[m.start(): m.start() + 4000]


def js_actions(js):
    """Every 'something.htm?action=xyz' the page's JavaScript calls."""
    return list(dict.fromkeys(re.findall(r"""["']([A-Za-z]+\.htm\?action=[A-Za-z0-9_]+)""", js)))


def options_in(html, name=None, prefer_selected=True):
    soup = BeautifulSoup(html or "", "lxml")
    scope = soup.find("select", attrs={"name": name}) if name else soup
    scope = scope or soup
    opts = [o for o in scope.find_all("option") if o.get("value") not in (None, "", "0", "-1")]
    chosen = [o for o in opts if o.has_attr("selected")]
    return [o.get("value") for o in (chosen if (prefer_selected and chosen) else opts)]


def parse_rows_fallback(html, ncols, colnames):
    """If APMI returns only table rows (no header), map cells by position."""
    soup = BeautifulSoup(html or "", "lxml")
    rows = [[c.get_text(" ", strip=True) for c in tr.find_all("td")] for tr in soup.find_all("tr")]
    rows = [r for r in rows if len(r) == ncols]
    if not rows:
        return None
    df = pd.DataFrame(rows, columns=colnames)
    for col in colnames[2:]:
        df[col] = df[col].map(clean_number)
    return df


def parse_perf_any(html):
    df = parse_performance(html)
    return df if df is not None else parse_rows_fallback(html, 12, list(PERF_COLS.values()))


def parse_turn_any(html):
    df = parse_turnover(html)
    return df if df is not None else parse_rows_fallback(html, 4, list(TURN_COLS.values()))


def date_formats(mid):
    y, m = mid.split("-")
    d = month_end(mid)[:2]
    # APMI's own script builds year-month-day with the month unpadded, e.g. 2026-8-31
    return [f"{y}-{int(m)}-{d}", f"{y}-{m}-{d}", month_end(mid)]


class ApmiReport:
    def __init__(self, session, page_url, html, kind):
        self.session, self.page_url, self.kind = session, page_url, kind
        self.js = js_sources(session, page_url, html)
        self.actions = js_actions(self.js)
        self.parse = parse_perf_any if kind == "perf" else parse_turn_any
        soup = BeautifulSoup(html, "lxml")
        radio_names = {i.get("name") for i in soup.find_all("input", attrs={"type": "radio"})}
        self.cat_field = next((n for n in radio_names if n and re.search(r"str[ae]t[ae]gy", n, re.I)), "strategyname")
        self.default_names = ia_names(self.parse(html))
        self.default_html = html
        self.samples = []
        self.report_action = None
        self.lists = {}
        self.locked = None
        self.log = []

    def url(self, action):
        return urljoin(self.page_url, action)

    def call(self, action, data, method="post"):
        time.sleep(DELAY_SECONDS)
        headers = dict(AJAX_HEADERS, Referer=self.page_url)
        if method == "post":
            r = self.session.post(self.url(action), data=data, headers=headers, timeout=90)
        else:
            r = self.session.get(self.url(action), params=data, headers=headers, timeout=90)
        if "charset" not in r.headers.get("Content-Type", "").lower():
            r.encoding = "utf-8"
        return r

    def report_actions(self):
        acts = [x for x in self.actions if not re.search(r"loadpms|loadianame|getia|iausing|iabypms|page$|pmsmenu|benchmark|chart|insight|excel|download", x, re.I)]
        preferred = [x for x in acts if re.search(r"report|data", x, re.I)]
        return list(dict.fromkeys(preferred + acts))

    def page_lists(self):
        """Provider / strategy IDs as they appear on the loaded (Equity) page."""
        return (options_in(self.default_html, "pmsProvideName"), options_in(self.default_html, "pmsInvAprochName"))

    def ia_lists(self, category, service):
        """Provider and investment-approach IDs for a category, fetched the way the page does."""
        key = (category, service)
        if key in self.lists:
            return self.lists[key]
        providers, ias = [], []
        pms_action = next((x for x in self.actions if re.search(r"loadpms", x, re.I)), None)
        ia_action = next((x for x in self.actions if re.search(r"loadianame|getia|iausing|iabypms", x, re.I)), None)
        page_prov = ",".join(self.page_lists()[0])
        try:
            if pms_action:
                r = self.call(pms_action, {"strategyname": category, "pmsProviderName": page_prov, "service": service})
                providers = options_in(r.text)
                self.samples.append(f"{pms_action} [{category}/{service}] -> {len(providers)} ids | {re.sub(r'\s+', ' ', r.text)[:200]}")
            if ia_action:
                joined = ",".join(providers)
                r = self.call(ia_action, {"strategyname": category, "pmsProvideName": joined, "pmsProviderName": joined, "service": service})
                ias = options_in(r.text)
                self.samples.append(f"{ia_action} [{category}/{service}] -> {len(ias)} ids | {re.sub(r'\s+', ' ', r.text)[:200]}")
        except requests.RequestException as e:
            self.log.append(f"list lookup failed: {e}")
        self.lists[key] = (providers, ias)
        return providers, ias

    def payload(self, category, service, mid, cfg):
        y, m = mid.split("-")
        providers, ias = self.page_lists() if cfg["ids"] == "page" else self.ia_lists(category, service)
        fields = [(self.cat_field, category), ("servicetype", service)]
        if cfg["style"] == "joined":
            fields += [("pmsProvideName", ",".join(providers)), ("pmsInvAprochName", ",".join(ias))]
        else:
            fields += [("pmsProvideName", v) for v in providers] + [("pmsInvAprochName", v) for v in ias]
        fields += [("fromMonth", str(int(m))), ("fromYears", y), ("asOnDate", date_formats(mid)[cfg.get("date", 0)])]
        return fields

    def run(self, category, service, mid, cfg):
        if cfg.get("bench_first"):
            bench = next((x for x in self.actions if re.search(r"benchmark", x, re.I)), None)
            if bench:
                self.call(bench, "")
        return self.call(self.report_action, self.payload(category, service, mid, cfg))

    def experiments(self):
        cfgs = []
        for bench in (False, True):
            for ids, cats in (("page", ["Equity"]), ("lists", ["Equity", "Debt"])):
                for style in ("repeated", "joined"):
                    for cat in cats:
                        cfgs.append((cat, {"ids": ids, "style": style, "bench_first": bench, "date": 0}))
        cfgs.append(("Equity", {"ids": "page", "style": "repeated", "bench_first": False, "date": 1}))
        return cfgs

    def probe(self, mid):
        acts = self.report_actions()
        if not acts:
            print(f"    {self.kind}: no report address found in APMI's JavaScript")
            return False
        self.report_action = acts[0]
        pp, pi = self.page_lists()
        print(f"    {self.kind}: report address {self.report_action}; page lists {len(pp)} providers, {len(pi)} strategies")
        working_equity = None
        for cat, cfg in self.experiments():
            try:
                r = self.run(cat, "D", mid, cfg)
            except requests.RequestException as e:
                self.log.append(f"{cat} {cfg}: error {e}")
                continue
            df = self.parse(r.text) if r.ok else None
            n = 0 if df is None else len(df)
            text = re.sub(r"\s+", " ", BeautifulSoup(r.text, "lxml").get_text(" "))
            self.log.append(f"{cat} {cfg}: HTTP {r.status_code}, {len(r.text)} bytes, rows={n} | {text[:120]} ... {text[-80:]}")
            if n and cat == "Equity" and working_equity is None:
                working_equity = cfg
            if n and cat == "Debt" and ia_names(df) != self.default_names:
                self.locked = cfg
                print(f"    {self.kind} probe OK: {cfg}")
                return True
        if working_equity:
            print(f"    {self.kind}: Equity works with {working_equity}, but Debt returned nothing")
        print(f"    {self.kind} probe failed")
        return False

    def fetch(self, category, service_code, mid):
        r = self.run(category, service_code, mid, self.locked)
        r.raise_for_status()
        return self.parse(r.text)

    def diagnostics(self):
        print(f"\n----- DIAGNOSTICS: {self.kind} -----")
        print(f"  category field: {self.cat_field}")
        print(f"  actions found in JavaScript: {self.actions}")
        print(f"  report actions tried: {self.report_actions()}")
        print(f"  lists: {[(k, len(v[0]), len(v[1])) for k, v in self.lists.items()]}")
        soup = BeautifulSoup(self.default_html, "lxml")
        for nm in ("pmsProvideName", "pmsInvAprochName"):
            sel = soup.find("select", attrs={"name": nm})
            if sel:
                opts = sel.find_all("option")
                print(f"  page select {nm}: multiple={sel.has_attr('multiple')} options={len(opts)} selected={sum(o.has_attr('selected') for o in opts)} parent-form={bool(sel.find_parent('form'))}")
        form = soup.find("form")
        if form:
            print("  named fields inside the form:", sorted({e.get('name') for e in form.find_all(['input', 'select', 'textarea']) if e.get('name')}))
        for line in self.samples[:6]:
            print("  list:", line)
        for line in self.log[:30]:
            print("  try:", line)
        for fn in ("getalldata", "getFormData", "onClickofSubmit", "onClickSubmit", "loadEquitydata", "strategytype", "IAbyService", "getIabyPMS", "getIaUsingPMS"):
            body = js_function(self.js, fn)
            if body:
                print(f"\n  JS {fn}:", re.sub(r"\s+", " ", body)[:2500])
        print("----- END DIAGNOSTICS -----\n")


SERVICE_CODES = {"Discretionary": "D", "Non Discretionary": "N"}


def scrape_month(perf, turn, mid, reference=None):
    frames, failures, seen = [], [], {}
    for category in CATEGORIES:
        for service in SERVICES:
            code = SERVICE_CODES[service]
            label = f"{mid} {category} / {service}"
            try:
                p = perf.fetch(category, code, mid)
                t = turn.fetch(category, code, mid) if turn.locked else None
                if (p is None or not len(p)) and (t is None or not len(t)):
                    print(f"  {label}: no strategies listed")
                    continue
                names = frozenset(ia_names(p if p is not None else t))
                if names and names in seen:
                    raise RuntimeError(f"same list as {seen[names]}; APMI ignored the filter")
                seen[names] = f"{category} / {service}"
                if reference is not None and p is not None and len(p):
                    ref = reference.get((category, service))
                    sig = round(float(pd.to_numeric(p.get("Return 1M (%)"), errors="coerce").fillna(0).sum()), 4)
                    if ref is not None and sig == ref:
                        raise RuntimeError("identical to the latest month; APMI ignored the date")
                combined = merge(p, t)
                combined.insert(0, "Service Type", service)
                combined.insert(0, "Category", category)
                frames.append(combined)
                print(f"  {label}: {len(combined)} strategies")
            except Exception as e:
                failures.append(f"{category} / {service}")
                print(f"  {label}: FAILED ({e})")
    return frames, failures


def signature(frames):
    out = {}
    for f in frames:
        key = (f["Category"].iloc[0], f["Service Type"].iloc[0])
        out[key] = round(float(pd.to_numeric(f.get("Return 1M (%)"), errors="coerce").fillna(0).sum()), 4)
    return out


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
def fetch(session, url):
    resp = session.get(url, timeout=60)
    resp.raise_for_status()
    return resp.text


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--site", default="docs", help="website folder to write into")
    ap.add_argument("--backfill", type=int, default=24, help="max older months to add this run (0 = none)")
    ap.add_argument("--earliest", default="2023-04", help="don't go back before this month (YYYY-MM)")
    args = ap.parse_args()

    session = requests.Session()
    session.headers.update(HEADERS)
    perf_default = fetch(session, PERF_URL)
    time.sleep(DELAY_SECONDS)
    turn_default = fetch(session, TURN_URL)

    latest_as_on = extract_as_on(perf_default)
    latest = month_id(latest_as_on)
    print(f"Latest month on APMI: {latest_as_on}")

    perf = ApmiReport(session, PERF_URL, perf_default, "perf")
    turn = ApmiReport(session, TURN_URL, turn_default, "turn")
    print(f"Performance JavaScript actions: {perf.actions}")
    print(f"Turnover JavaScript actions:    {turn.actions}")

    print("\nChecking how APMI answers requests:")
    perf_ok = perf.probe(latest)
    turn_ok = turn.probe(latest)
    if not perf_ok:
        perf.diagnostics()
    if not turn_ok:
        turn.diagnostics()

    frames, failures = [], []
    if perf_ok:
        if not turn_ok:
            print("Turnover requests aren't working yet; returns and AUM will be saved without turnover.")
        print(f"\nLatest month {latest}:")
        frames, failures = scrape_month(perf, turn, latest)
    if not frames:
        print("Category requests aren't working yet; saving APMI's default Equity view so the site still updates.")
        data = merge(parse_performance(perf_default), parse_turnover(turn_default))
        data.insert(0, "Service Type", "Default view")
        data.insert(0, "Category", "Equity")
        failures = ["default-only"]
    else:
        data = pd.concat(frames, ignore_index=True)
        if not turn_ok:
            failures.append("turnover")
    data.insert(0, "As On", latest_as_on)

    payload = month_payload(data, latest_as_on, failures)
    write_json(os.path.join(args.site, "data.json"), payload)
    write_json(os.path.join(args.site, "history", f"{latest}.json"), payload)
    with pd.ExcelWriter(os.path.join(args.site, "apmi_pms_data.xlsx"), engine="openpyxl") as xl:
        data.to_excel(xl, sheet_name="Latest month", index=False)
        xl.sheets["Latest month"].freeze_panes = "F2"
        xl.sheets["Latest month"].auto_filter.ref = xl.sheets["Latest month"].dimensions
    print(f"Saved latest month: {len(data)} strategies")

    if frames and args.backfill > 0:
        ref = signature(frames)
        done, empty_streak, mid = 0, 0, prev_month(latest)
        while mid >= args.earliest and done < args.backfill:
            if not history_is_complete(args.site, mid):
                print(f"\nHistory {mid}:")
                f_hist, fail_hist = scrape_month(perf, turn, mid, reference=ref)
                if f_hist:
                    empty_streak = 0
                    hist = pd.concat(f_hist, ignore_index=True)
                    hist.insert(0, "As On", month_end(mid))
                    if not turn_ok:
                        fail_hist.append("turnover")
                    write_json(os.path.join(args.site, "history", f"{mid}.json"), month_payload(hist, month_end(mid), fail_hist))
                    done += 1
                    print(f"  saved {mid}: {len(hist)} strategies")
                else:
                    empty_streak += 1
                    if empty_streak >= 2:
                        print("  No data for two months in a row; stopping history here.")
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
