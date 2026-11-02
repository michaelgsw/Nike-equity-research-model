"""
Nike, Inc. (NYSE: NKE) - SEC EDGAR data pipeline.

Downloads the XBRL files for Nike's last five 10-Ks and the most recent 10-Q,
extracts the three primary statements (as presented by Nike) plus segment
revenue / EBIT, and writes tidy CSVs with a source for every number.

Usage:
    python data_pipeline.py              # download (cached) + build CSVs
    python data_pipeline.py --refresh    # ignore the cache and re-download

Set SEC_USER_AGENT first (name + email, as the SEC requires), e.g. in PowerShell:
    $env:SEC_USER_AGENT = "Jane Doe jane@example.com"

Outputs (data/clean/):
    filings.csv               the filings used (form, accession, dates, URL)
    statements_long.csv       every IS/BS/CF fact from every filing (one row per filing x line x period)
    statements_selected.csv   each statement x period taken whole from the most recent filing that reports it
    statements_wide_<IS|BS|CF>.csv   preview tables, USD millions, fiscal years as columns
    segments_long.csv         segment revenue / EBIT / costs by geography, channel and product, every filing
    segments_selected.csv     segment facts per period from the most recent filing that reports it
    segments_wide.csv         preview: revenue by segment x channel and EBIT by segment
    supplementary.csv         note data: dividends per share, debt maturities and bond table, tax rate, shares outstanding
    restatements.csv          values revised, or lines added/removed, between an original and a later filing
    checks.csv                tie-out checks (BS balances, subtotals, cash roll-forward, segments vs. totals)

Raw downloads are cached, never edited, in data/raw/<accession>/.

Conventions:
    - Nike's fiscal year ends May 31. FY2026 = year ended May 31, 2026.
    - value_usd_m keeps the XBRL sign (e.g. cost of sales is positive).
      value_presented flips the sign where Nike's statement shows the line negated
      (e.g. "Repurchase of common stock" is shown in parentheses on the cash flow statement).
"""

import argparse
import json
import os
import sys
import time
import xml.etree.ElementTree as ET
from datetime import date, timedelta

import pandas as pd
import requests

# --- Configuration -----------------------------------------------------------

CIK = "320187"
N_ANNUAL = 5  # number of 10-Ks to pull
# SEC asks for a descriptive User-Agent with contact details, e.g. "Jane Doe jane@example.com".
# It is read from the SEC_USER_AGENT environment variable so no email address lives in the code.
USER_AGENT = os.environ.get("SEC_USER_AGENT", "")
REQUEST_PAUSE_S = 0.2  # SEC fair-access limit is 10 requests/second; stay well below it

HERE = os.path.dirname(os.path.abspath(__file__))
RAW_DIR = os.path.join(HERE, "data", "raw")
CLEAN_DIR = os.path.join(HERE, "data", "clean")

NS = {
    "xbrli": "http://www.xbrl.org/2003/instance",
    "link": "http://www.xbrl.org/2003/linkbase",
    "xbrldi": "http://xbrl.org/2006/xbrldi",
}
XLINK = "{http://www.w3.org/1999/xlink}"

# Which reports in FilingSummary.xml are the statements we want (matched on the report's short name).
STATEMENT_MATCHERS = {
    "IS": lambda n: "statements of income" in n and "comprehensive" not in n,
    "BS": lambda n: "balance sheets" in n and "parenthetical" not in n,
    "CF": lambda n: "statements of cash flows" in n,
}
# Segment detail reports: segment table and revenue disaggregation table.
SEGMENT_MATCHER = lambda n: ("information by operating segments" in n) or (
    n.startswith("revenues") and ("disaggregation" in n or n == "revenues (details)")
)
CLASS_AXIS = "us-gaap:StatementClassOfStockAxis"
SEGMENT_AXES = {
    "us-gaap:StatementBusinessSegmentsAxis": "axis_segment",
    "us-gaap:SubsegmentsAxis": "axis_subsegment",
    "us-gaap:ContractWithCustomerSalesChannelAxis": "axis_channel",
    "srt:ProductOrServiceAxis": "axis_product",
    "srt:ConsolidationItemsAxis": "axis_consolidation",
}


# --- HTTP with cache ----------------------------------------------------------

session = requests.Session()
session.headers.update({"User-Agent": USER_AGENT, "Accept-Encoding": "gzip, deflate"})


def require_user_agent():
    if not USER_AGENT.strip():
        raise RuntimeError('Set the SEC_USER_AGENT environment variable to "Your Name your@email.com" '
                           "(SEC fair-access policy) before downloading from EDGAR.")


def fetch(url, cache_path, refresh=False):
    """Download url to cache_path (unless cached) and return the file path."""
    if os.path.exists(cache_path) and not refresh:
        return cache_path
    require_user_agent()
    os.makedirs(os.path.dirname(cache_path), exist_ok=True)
    for attempt in range(4):
        time.sleep(REQUEST_PAUSE_S)
        resp = session.get(url, timeout=60)
        if resp.status_code == 429:  # rate limited: back off and retry
            time.sleep(2 ** attempt)
            continue
        resp.raise_for_status()
        with open(cache_path, "wb") as f:
            f.write(resp.content)
        return cache_path
    raise RuntimeError(f"Rate-limited repeatedly fetching {url}")


# --- Filing discovery ---------------------------------------------------------

def list_filings(refresh):
    """Return the last N_ANNUAL 10-Ks plus the latest 10-Q filed after the latest 10-K."""
    path = fetch(f"https://data.sec.gov/submissions/CIK{CIK.zfill(10)}.json",
                 os.path.join(RAW_DIR, "submissions.json"), refresh=True)  # always refresh the filing list
    with open(path) as f:
        recent = json.load(f)["filings"]["recent"]
    df = pd.DataFrame({k: recent[k] for k in ["form", "accessionNumber", "filingDate", "reportDate", "primaryDocument"]})
    tenk = df[df.form == "10-K"].sort_values("reportDate", ascending=False).head(N_ANNUAL)
    tenq = df[(df.form == "10-Q") & (df.filingDate > tenk.filingDate.max())].sort_values("reportDate").tail(1)
    out = pd.concat([tenk, tenq]).sort_values("reportDate").reset_index(drop=True)
    out["folder_url"] = out.accessionNumber.map(
        lambda a: f"https://www.sec.gov/Archives/edgar/data/{CIK}/{a.replace('-', '')}")
    out["document_url"] = out.folder_url + "/" + out.primaryDocument
    return out


def download_filing(row, refresh):
    """Download FilingSummary, XBRL instance, presentation and label linkbases for one filing."""
    acc = row.accessionNumber
    folder = os.path.join(RAW_DIR, acc)
    index_path = fetch(row.folder_url + "/index.json", os.path.join(folder, "index.json"), refresh)
    with open(index_path) as f:
        names = [i["name"] for i in json.load(f)["directory"]["item"]]
    files = {
        "summary": "FilingSummary.xml",
        "instance": next(n for n in names if n.endswith("_htm.xml")),
        "pre": next(n for n in names if n.endswith("_pre.xml")),
        "lab": next(n for n in names if n.endswith("_lab.xml")),
    }
    return {k: fetch(f"{row.folder_url}/{v}", os.path.join(folder, v), refresh) for k, v in files.items()}


# --- XBRL parsing -------------------------------------------------------------

def fiscal_label(end, months):
    """Nike FY ends May 31: a period ending Jun-Dec belongs to the next fiscal year."""
    fy = end.year if end.month <= 5 else end.year + 1
    if months is None:  # balance sheet date
        return f"FY{fy}" if (end.month, end.day) == (5, 31) else f"Q{quarter_of(end)} FY{fy}"
    if months == 12:
        return f"FY{fy}"
    if months == 3:
        return f"Q{quarter_of(end)} FY{fy}"
    return f"{months}M FY{fy}"


def quarter_of(end):
    return {8: 1, 9: 1, 11: 2, 12: 2, 2: 3, 3: 3, 5: 4, 6: 4}.get(end.month, 0)


def parse_instance(path):
    """Return a DataFrame of all numeric facts with their period and dimensions."""
    root = ET.parse(path).getroot()
    contexts = {}
    for c in root.findall("xbrli:context", NS):
        p = c.find("xbrli:period", NS)
        inst = p.find("xbrli:instant", NS)
        if inst is not None:
            start, end = None, date.fromisoformat(inst.text)
        else:
            start = date.fromisoformat(p.find("xbrli:startDate", NS).text)
            end = date.fromisoformat(p.find("xbrli:endDate", NS).text)
        dims = {m.get("dimension"): m.text for m in c.iter(f"{{{NS['xbrldi']}}}explicitMember")}
        contexts[c.get("id")] = (start, end, dims)

    # e.g. "iso4217:USD", "shares", or "iso4217:USD/shares" for per-share values
    units = {u.get("id"): "/".join(m.text.strip() for m in u.iter(f"{{{NS['xbrli']}}}measure"))
             for u in root.findall("xbrli:unit", NS)}
    rows = []
    for el in root:
        ctx = el.get("contextRef")
        if ctx is None or el.get("unitRef") is None or el.text is None:
            continue
        prefix_uri, local = el.tag[1:].split("}")
        prefix = "us-gaap" if "fasb.org/us-gaap" in prefix_uri else ("ifrs-full" if "xbrl.ifrs.org" in prefix_uri else ("nke" if "nike.com" in prefix_uri else prefix_uri.rsplit("/", 2)[-2]))
        start, end, dims = contexts[ctx]
        rows.append({
            "concept": f"{prefix}_{local}", "value": float(el.text), "unit": units[el.get("unitRef")],
            "start": start, "end": end, "dims": dims, "decimals": el.get("decimals"),
        })
    df = pd.DataFrame(rows)
    # Inline XBRL repeats a fact wherever it appears in the document; keep one copy.
    df["dims_key"] = df.dims.map(lambda d: tuple(sorted(d.items())))
    return df.drop_duplicates(subset=["concept", "start", "end", "dims_key", "value"]).reset_index(drop=True)


def parse_labels(path):
    """Map (concept, label role) -> label text."""
    root = ET.parse(path).getroot()
    labels = {}
    for lnk in root.findall("link:labelLink", NS):
        loc = {l.get(XLINK + "label"): l.get(XLINK + "href").split("#")[1] for l in lnk.findall("link:loc", NS)}
        lab = {}
        for l in lnk.findall("link:label", NS):
            lab.setdefault(l.get(XLINK + "label"), []).append((l.get(XLINK + "role"), l.text))
        for arc in lnk.findall("link:labelArc", NS):
            for role, text in lab.get(arc.get(XLINK + "to"), []):
                labels[(loc[arc.get(XLINK + "from")], role.split("/")[-1])] = text
    return labels


def statement_reports(summary_path):
    """Map statement code -> (presentation role URI, report short name) from FilingSummary.xml."""
    root = ET.parse(summary_path).getroot()
    found = {}
    for r in root.iter("Report"):
        name = (r.findtext("ShortName") or "").strip()
        role = r.findtext("Role")
        for code, match in STATEMENT_MATCHERS.items():
            if match(name.lower()) and code not in found:
                found[code] = (role, name)
        if SEGMENT_MATCHER(name.lower()):
            found.setdefault("SEG", []).append((role, name))
    return found


def presentation_lines(pre_path, role):
    """Walk the presentation tree of one role; return ordered line items (concept, preferred label, depth)."""
    root = ET.parse(pre_path).getroot()
    for pl in root.findall("link:presentationLink", NS):
        if pl.get(XLINK + "role") != role:
            continue
        loc = {l.get(XLINK + "label"): l.get(XLINK + "href").split("#")[1] for l in pl.findall("link:loc", NS)}
        children, targets = {}, set()
        for a in pl.findall("link:presentationArc", NS):
            pref = (a.get("preferredLabel") or "label").split("/")[-1]
            children.setdefault(a.get(XLINK + "from"), []).append((float(a.get("order", 0)), a.get(XLINK + "to"), pref))
            targets.add(a.get(XLINK + "to"))
        lines = []

        def walk(node, depth, pref):
            concept = loc[node]
            if not concept.endswith(("Abstract", "Axis", "Domain", "Member", "Table", "LineItems")):
                lines.append((concept, pref, depth))
            for _, child, child_pref in sorted(children.get(node, [])):
                walk(child, depth + 1, child_pref)

        for r in [n for n in children if n not in targets]:
            walk(r, 0, "label")
        return lines
    return []


# --- Extraction ---------------------------------------------------------------

def extract_statements(filing, paths):
    facts = parse_instance(paths["instance"])
    labels = parse_labels(paths["lab"])
    reports = statement_reports(paths["summary"])
    plain = facts[facts.dims.map(len) == 0]  # statement face values carry no dimensions
    # Balance sheet dates = dates with a Total assets figure (drops equity-statement balances at older dates).
    bs_dates = set(plain[plain.concept == "us-gaap_Assets"].end)
    rows = []
    for code in ["IS", "BS", "CF"]:
        role, report_name = reports[code]
        for order, (concept, pref, depth) in enumerate(presentation_lines(paths["pre"], role)):
            label = labels.get((concept, pref)) or labels.get((concept, "label"), concept)
            f = plain[plain.concept == concept]
            if code == "CF" and pref in ("periodStartLabel", "periodEndLabel"):
                # Cash roll-forward lines are instants; attach them to each CF duration period.
                cf_periods = plain[(plain.concept == "us-gaap_NetCashProvidedByUsedInOperatingActivities")][["start", "end"]]
                for s, e in cf_periods.drop_duplicates().itertuples(index=False):
                    on = s - timedelta(days=1) if pref == "periodStartLabel" else e
                    hit = f[f.end == on]
                    if len(hit):
                        rows.append(make_row(filing, code, report_name, order, depth, concept, label, pref, hit.iloc[0], s, e))
                continue
            if f.empty and code == "BS":
                # Nike reports common stock only by share class (Class A / Class B), i.e. as dimensional facts.
                by_class = facts[(facts.concept == concept) & facts.dims.map(lambda d: set(d) == {CLASS_AXIS})]
                for i, member in enumerate(sorted(by_class.dims.map(lambda d: d[CLASS_AXIS]).unique())):
                    member_label = labels.get((member.replace(":", "_"), "label"), member)
                    for _, fact in by_class[by_class.dims.map(lambda d: d[CLASS_AXIS] == member)].iterrows():
                        if fact.end in bs_dates:
                            rows.append(make_row(filing, code, report_name, order + i / 10, depth, f"{concept}[{member}]",
                                                 f"{label} - {member_label}", pref, fact, None, fact.end))
                if len(by_class):
                    continue
            if f.empty:
                rows.append(make_row(filing, code, report_name, order, depth, concept, label, pref, None, None, None))
                continue
            for _, fact in f.iterrows():
                if code == "BS" and (fact.start is not None or fact.end not in bs_dates):
                    continue
                if code in ("IS", "CF") and fact.start is None:
                    continue
                rows.append(make_row(filing, code, report_name, order, depth, concept, label, pref, fact, fact.start, fact.end))
    return pd.DataFrame(rows), facts, labels, reports


def make_row(filing, code, report_name, order, depth, concept, label, pref, fact, start, end):
    months = None if start is None else round((end - start).days / 30.4)
    scale = 1e6 if (fact is not None and fact.unit in ("iso4217:USD", "shares")) else 1.0
    value = None if fact is None else fact.value / scale
    negated = pref.startswith("negated")
    return {
        "statement": code, "line_order": order, "depth": depth, "concept": concept, "label": label,
        # line_id separates the two cash roll-forward lines, which share one XBRL concept
        "line_id": concept + {"periodStartLabel": "[start]", "periodEndLabel": "[end]"}.get(pref, ""),
        "period": None if end is None else fiscal_label(end, months),
        "period_start": start, "period_end": end, "months": months,
        "value_usd_m": value,
        "value_presented": None if value is None else (-value if negated else value),
        "unit": None if fact is None else ("USD m" if fact.unit == "iso4217:USD" else "shares m" if fact.unit == "shares" else fact.unit),
        "form": filing.form, "accession": filing.accessionNumber, "filing_date": filing.filingDate,
        "source": f"{filing.form} {filing.accessionNumber} | {report_name} | XBRL {concept.replace('_', ':', 1)}",
        "document_url": filing.document_url,
    }


def extract_segments(filing, facts, labels, reports, pre_path):
    """Dimensional revenue / EBIT facts from the segment and revenue-disaggregation notes."""
    concepts = set()
    for role, _ in reports.get("SEG", []):
        concepts |= {c for c, _, _ in presentation_lines(pre_path, role)}
    seg = facts[facts.concept.isin(concepts) & (facts.dims.map(len) > 0) & facts.start.notna()].copy()
    # keep only facts whose dimensions are all segment-type axes
    seg = seg[seg.dims.map(lambda d: set(d) <= set(SEGMENT_AXES))]
    for axis, col in SEGMENT_AXES.items():
        seg[col] = seg.dims.map(lambda d: (d.get(axis) or "").split(":")[-1])
    seg["months"] = [round((e - s).days / 30.4) for s, e in zip(seg.start, seg.end)]
    seg["period"] = [fiscal_label(e, m) for e, m in zip(seg.end, seg.months)]
    seg["segment"] = seg.apply(canonical_segment, axis=1)
    seg["channel"] = seg.axis_channel.map(canonical_channel)
    seg["product"] = seg.axis_product.str.replace("Member", "", regex=False).replace("", "All")
    seg["label"] = seg.concept.map(lambda c: labels.get((c, "label"), c))
    seg["value_usd_m"] = seg.value / 1e6
    seg["form"], seg["accession"], seg["filing_date"] = filing.form, filing.accessionNumber, filing.filingDate
    seg["source"] = [f"{filing.form} {filing.accessionNumber} | segment/revenue note | XBRL {c.replace('_', ':', 1)} "
                     f"[{dims_text(d)}]" for c, d in zip(seg.concept, seg.dims)]
    cols = ["concept", "label", "period", "start", "end", "months", "segment", "channel", "product",
            "axis_segment", "axis_subsegment", "axis_channel", "axis_product", "axis_consolidation",
            "value_usd_m", "form", "accession", "filing_date", "source"]
    return seg[cols].rename(columns={"start": "period_start", "end": "period_end"})


SUPPLEMENTARY = {  # note disclosures the model uses: concept -> description
    "us-gaap_CommonStockDividendsPerShareDeclared": "Dividends declared per common share (USD)",
    "us-gaap_LongTermDebtMaturitiesRepaymentsOfPrincipalInNextTwelveMonths": "Debt maturities - year 1",
    "us-gaap_LongTermDebtMaturitiesRepaymentsOfPrincipalInYearTwo": "Debt maturities - year 2",
    "us-gaap_LongTermDebtMaturitiesRepaymentsOfPrincipalInYearThree": "Debt maturities - year 3",
    "us-gaap_LongTermDebtMaturitiesRepaymentsOfPrincipalInYearFour": "Debt maturities - year 4",
    "us-gaap_LongTermDebtMaturitiesRepaymentsOfPrincipalInYearFive": "Debt maturities - year 5",
    "us-gaap_EffectiveIncomeTaxRateContinuingOperations": "Effective tax rate",
    "dei_EntityCommonStockSharesOutstanding": "Shares outstanding (cover page)",
    "us-gaap_RestructuringCharges": "Severance and restructuring charges",
}
LOCATION_AXIS = "us-gaap:IncomeStatementLocationAxis"  # splits restructuring between cost of sales and overhead
DEBT_AXIS = "us-gaap:DebtInstrumentAxis"


def extract_supplementary(filing, facts):
    """Selected note facts (dividends per share, debt maturities, tax rate, shares outstanding)
    plus the bond-by-bond debt table (face amount, coupon, carrying amount)."""
    rows = []
    for concept, desc in SUPPLEMENTARY.items():
        f = facts[facts.concept == concept]
        if concept.startswith("dei_"):  # cover page reports shares by class
            f = f[f.dims.map(lambda d: set(d) <= {CLASS_AXIS})]
        elif concept == "us-gaap_RestructuringCharges":  # keep the income statement location split
            f = f[f.dims.map(lambda d: set(d) == {LOCATION_AXIS})]
        else:
            f = f[f.dims.map(len) == 0]
        for r in f.itertuples():
            months = None if r.start is None else round((r.end - r.start).days / 30.4)
            rows.append({"item": desc, "concept": concept, "instrument": dims_text(r.dims),
                         "period": fiscal_label(r.end, months) if concept != "dei_EntityCommonStockSharesOutstanding" else str(r.end),
                         "period_end": r.end, "value": r.value / (1e6 if r.unit in ("iso4217:USD", "shares") else 1),
                         "unit": "USD m" if r.unit == "iso4217:USD" else "shares m" if r.unit == "shares" else r.unit})
    debt = facts[facts.dims.map(lambda d: DEBT_AXIS in d) & facts.concept.isin(
        ["us-gaap_DebtInstrumentFaceAmount", "us-gaap_DebtInstrumentInterestRateStatedPercentage", "us-gaap_LongTermDebt"])]
    for r in debt.itertuples():
        if r.start is None and r.end.month == 5 and r.end.day == 31:
            rows.append({"item": {"us-gaap_DebtInstrumentFaceAmount": "Debt face amount",
                                  "us-gaap_DebtInstrumentInterestRateStatedPercentage": "Debt coupon rate",
                                  "us-gaap_LongTermDebt": "Debt carrying amount"}[r.concept],
                         "concept": r.concept, "instrument": r.dims[DEBT_AXIS].split(":")[-1],
                         "period": fiscal_label(r.end, None), "period_end": r.end,
                         "value": r.value / (1e6 if r.unit == "iso4217:USD" else 1),
                         "unit": "USD m" if r.unit == "iso4217:USD" else r.unit})
    out = pd.DataFrame(rows)
    out["form"], out["accession"], out["filing_date"] = filing.form, filing.accessionNumber, filing.filingDate
    out["source"] = [f"{filing.form} {filing.accessionNumber} | notes | XBRL {c.replace('_', ':', 1)}" for c in out.concept]
    return out.drop_duplicates(subset=["concept", "instrument", "period", "value"])


def dims_text(dims):
    return "; ".join(f"{axis.split(':')[-1]}={member.split(':')[-1]}" for axis, member in sorted(dims.items()))


def canonical_segment(r):
    names = {"NorthAmerica": "North America", "EuropeMiddleEastAndAfrica": "EMEA", "GreaterChina": "Greater China",
             "AsiaPacificAndLatinAmerica": "APLA", "GlobalBrandDivisions": "Global Brand Divisions",
             "Converse": "Converse", "Corporate": "Corporate", "NIKEBrand": "NIKE Brand"}
    # Global Brand Divisions is tagged as the reconciling item (under NIKE Brand from FY2025, standalone before).
    if r.axis_consolidation == "MaterialReconcilingItemsMember" and not r.axis_subsegment and r.axis_segment in ("", "NIKEBrandMember"):
        return "Global Brand Divisions"
    for raw in (r.axis_subsegment, r.axis_segment):
        for key, name in names.items():
            if raw.startswith(key):
                return name
    if r.axis_consolidation.startswith("Corporate"):
        return "Corporate"
    return "Total NIKE, Inc." if not r.axis_segment and not r.axis_subsegment else r.axis_subsegment or r.axis_segment


def canonical_channel(raw):
    if not raw:
        return "All"
    if "Wholesale" in raw:
        return "Wholesale"
    if "DirectlyToConsumer" in raw or "NIKEDirect" in raw:
        return "NIKE Direct"
    if "Other" in raw:
        return "Other"
    return raw


# --- Selection, restatements, checks -------------------------------------------

def select_latest(long):
    """For each statement x period, keep all lines from the single most recent filing that reports it.

    Taking a whole statement from one filing (rather than line by line) keeps each period internally
    consistent if Nike renames or regroups a line between filings.
    """
    has = long.dropna(subset=["value_usd_m"])
    # Annual periods come from 10-Ks; the 10-Q contributes only its own quarter and balance sheet date.
    has = has[(has.form == "10-K") | ~has.period.str.match(r"^FY\d{4}$")]
    latest = has.groupby(["statement", "period"]).filing_date.max().rename("latest_date").reset_index()
    has = has.merge(latest, on=["statement", "period"])
    return has[has.filing_date == has.latest_date].drop(columns="latest_date").reset_index(drop=True)


def find_restatements(long):
    """Values that changed between filings, and lines that appear in only one of the two filings."""
    has = long.dropna(subset=["value_usd_m"])
    out = []
    for (stmt, period), grp in has.groupby(["statement", "period"]):
        accs = grp.sort_values("filing_date").accession.unique()
        if len(accs) < 2:
            continue
        first, last = grp[grp.accession == accs[0]], grp[grp.accession == accs[-1]]
        merged = first.merge(last, on="line_id", how="outer", suffixes=("_orig", "_latest"))
        for r in merged.itertuples():
            if pd.isna(r.value_usd_m_orig) or pd.isna(r.value_usd_m_latest) or abs(r.value_usd_m_orig - r.value_usd_m_latest) > 1e-9:
                out.append({"statement": stmt, "period": period, "line_id": r.line_id,
                            "label": r.label_latest if isinstance(r.label_latest, str) else r.label_orig,
                            "original_value": r.value_usd_m_orig, "original_filing": accs[0],
                            "latest_value": r.value_usd_m_latest, "latest_filing": accs[-1],
                            "note": "line not in original filing" if pd.isna(r.value_usd_m_orig)
                            else "line not in latest filing" if pd.isna(r.value_usd_m_latest) else "value revised"})
    return pd.DataFrame(out)


def run_checks(sel):
    """Arithmetic tie-outs on the selected statements (tolerance 0.5 = rounding to USD millions)."""
    checks = []

    def add(name, period, lhs, rhs, info_only=False):
        ok = None if info_only or lhs is None or rhs is None else abs(lhs - rhs) < 0.5
        checks.append({"check": name, "period": period, "lhs": lhs, "rhs": rhs,
                       "difference": None if lhs is None or rhs is None else round(lhs - rhs, 3), "pass": ok})

    for (stmt, p), lines in sel.groupby(["statement", "period"]):
        lines = lines.sort_values("line_order")
        pos = {c: o for c, o in zip(lines.concept, lines.line_order)}

        def val(concept, which=0):
            hit = lines[lines.concept == concept]
            return None if len(hit) <= which else hit.iloc[which].value_usd_m

        def between(start, end):
            """Sum of presented values of lines strictly between two concepts (start=None: from the top)."""
            if end not in pos or (start is not None and start not in pos):
                return None
            lo = -1 if start is None else pos[start]
            return lines[(lines.line_order > lo) & (lines.line_order < pos[end])].value_presented.sum()

        def diff(a, b):
            return None if a is None or b is None else a - b

        if stmt == "BS":
            ta, tle = val("us-gaap_Assets"), val("us-gaap_LiabilitiesAndStockholdersEquity")
            ca, cl, se = val("us-gaap_AssetsCurrent"), val("us-gaap_LiabilitiesCurrent"), val("us-gaap_StockholdersEquity")
            add("BS: total assets = total liabilities and equity", p, ta, tle)
            add("BS: current asset lines sum to total current assets", p, between(None, "us-gaap_AssetsCurrent"), ca)
            nca = between("us-gaap_AssetsCurrent", "us-gaap_Assets")
            add("BS: current + non-current asset lines = total assets", p, None if None in (ca, nca) else ca + nca, ta)
            add("BS: current liability lines sum to total current liabilities", p, between("us-gaap_Assets", "us-gaap_LiabilitiesCurrent"), cl)
            add("BS: equity lines sum to total shareholders' equity", p,
                between("us-gaap_TemporaryEquityCarryingAmountAttributableToParent", "us-gaap_StockholdersEquity"), se)
            rest = between("us-gaap_LiabilitiesCurrent", "us-gaap_LiabilitiesAndStockholdersEquity")
            # 'rest' covers non-current liabilities, preferred stock, equity lines AND the equity subtotal, so remove the subtotal once
            add("BS: liability and equity lines = total liabilities and equity", p, None if None in (cl, rest, se) else cl + rest - se, tle)
        elif stmt == "IS":
            gp, sga = val("us-gaap_GrossProfit"), val("us-gaap_SellingGeneralAndAdministrativeExpense")
            pbt_c = "us-gaap_IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest"
            add("IS: revenue - cost of sales = gross profit", p,
                diff(val("us-gaap_RevenueFromContractWithCustomerExcludingAssessedTax"), val("us-gaap_CostOfGoodsAndServicesSold")), gp)
            add("IS: demand creation + operating overhead = total SG&A", p, between("us-gaap_GrossProfit", "us-gaap_SellingGeneralAndAdministrativeExpense"), sga)
            below = between("us-gaap_SellingGeneralAndAdministrativeExpense", pbt_c)  # interest and other (presented as expense)
            add("IS: gross profit - SG&A - interest - other = pre-tax income", p, None if None in (gp, sga, below) else gp - sga - below, val(pbt_c))
            add("IS: pre-tax income - tax = net income", p, diff(val(pbt_c), val("us-gaap_IncomeTaxExpenseBenefit")), val("us-gaap_NetIncomeLoss"))
        elif stmt == "CF":
            cfo, cfi, cff = ("us-gaap_NetCashProvidedByUsedIn" + x + "Activities" for x in ("Operating", "Investing", "Financing"))
            fx = "us-gaap_EffectOfExchangeRateOnCashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents"
            chg = "us-gaap_CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsPeriodIncreaseDecreaseIncludingExchangeRateEffect"
            cash = "us-gaap_CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents"
            add("CF: operating lines sum to cash from operations", p, between(None, cfo), val(cfo))
            add("CF: investing lines sum to cash from investing", p, between(cfo, cfi), val(cfi))
            add("CF: financing lines sum to cash from financing", p, between(cfi, cff), val(cff))
            parts = [val(c) for c in (cfo, cfi, cff, fx)]
            add("CF: CFO + CFI + CFF + FX = net change in cash", p, None if None in parts else sum(parts), val(chg))
            begin, end = val(cash, which=0), val(cash, which=1)
            add("CF: beginning cash + net change = ending cash", p, None if None in (begin, val(chg)) else begin + val(chg), end)
            bs_cash = sel[(sel.statement == "BS") & (sel.period == p) & (sel.concept == "us-gaap_CashAndCashEquivalentsAtCarryingValue")].value_usd_m
            # Not applicable (pass = blank) for periods without a balance sheet in the selected data
            add("CF ending cash = BS cash and equivalents", p, end, None if bs_cash.empty else bs_cash.iloc[0])
    return pd.DataFrame(checks)


def select_segments(seg):
    """Segment facts for each period from the most recent filing that reports that period (10-Ks for FYs)."""
    seg = seg[(seg.form == "10-K") | ~seg.period.str.match(r"^FY\d{4}$")]
    latest = seg.groupby("period").filing_date.transform("max")
    return seg[seg.filing_date == latest].reset_index(drop=True)


SEGMENT_REVENUE = "us-gaap_RevenueFromContractWithCustomerExcludingAssessedTax"
GEOS = ["North America", "EMEA", "Greater China", "APLA"]


def segment_checks(seg_sel, sel):
    """Geographies + Global Brand Divisions = NIKE Brand; + Converse + Corporate = total = income statement revenue;
    wholesale + NIKE Direct + other = segment revenue."""
    rev = seg_sel[(seg_sel.concept == SEGMENT_REVENUE) & (seg_sel["product"] == "All")]
    checks = []
    for p, r in rev.groupby("period"):
        tot = r[r.channel == "All"].drop_duplicates("segment").set_index("segment").value_usd_m
        is_rev = sel[(sel.statement == "IS") & (sel.period == p) & (sel.concept == SEGMENT_REVENUE)].value_usd_m
        brand_parts = tot.reindex(GEOS + ["Global Brand Divisions"]).fillna(0).sum()
        company = tot.reindex(["NIKE Brand", "Converse", "Corporate"]).fillna(0).sum()
        pairs = [("SEG: geographies + Global Brand Divisions = NIKE Brand revenue", brand_parts, tot.get("NIKE Brand")),
                 ("SEG: NIKE Brand + Converse + Corporate = income statement revenue", company,
                  None if is_rev.empty else is_rev.iloc[0])]
        # Segment EBIT (geographies + GBD + Converse + Corporate) = pre-tax income - net interest income
        ebit = seg_sel[(seg_sel.concept == "nke_EarningsBeforeInterestAndTaxes") & (seg_sel.period == p) & (seg_sel["product"] == "All")]
        ebit = ebit.drop_duplicates("segment").set_index("segment").value_usd_m
        is_p = sel[(sel.statement == "IS") & (sel.period == p)].set_index("concept").value_usd_m
        pbt = is_p.get("us-gaap_IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest")
        interest = is_p.get("us-gaap_InterestIncomeExpenseNonoperatingNet")  # XBRL sign: positive = net income
        if len(ebit) and pbt is not None and interest is not None:
            pairs.append(("SEG: segment EBIT incl. Corporate = pre-tax income - net interest income",
                          ebit.reindex(GEOS + ["Global Brand Divisions", "Converse", "Corporate"]).fillna(0).sum(), pbt - interest))
        for s in GEOS + ["Converse", "NIKE Brand"]:
            ch = r[(r.segment == s) & (r.channel != "All")].drop_duplicates("channel").value_usd_m.sum()
            pairs.append((f"SEG: {s} wholesale + NIKE Direct + other = segment revenue", ch, tot.get(s)))
        for name, lhs, rhs in pairs:
            checks.append({"check": name, "period": p, "lhs": lhs, "rhs": rhs,
                           "difference": None if rhs is None else round(lhs - rhs, 3),
                           "pass": None if rhs is None else abs(lhs - rhs) < 0.5})
    return pd.DataFrame(checks)


def segments_wide(seg_sel, periods):
    """Preview: revenue by segment and channel, and EBIT by segment, USD millions."""
    s = seg_sel[seg_sel.period.isin(periods) & (seg_sel["product"] == "All")]
    order = GEOS + ["Global Brand Divisions", "NIKE Brand", "Converse", "Corporate", "Total NIKE, Inc."]
    out = []
    for metric, concept, channels in [("Revenue", SEGMENT_REVENUE, ["All", "Wholesale", "NIKE Direct", "Other"]),
                                      ("EBIT", "nke_EarningsBeforeInterestAndTaxes", ["All"])]:
        m = s[(s.concept == concept) & s.channel.isin(channels)]
        t = m.pivot_table(index=["segment", "channel"], columns="period", values="value_usd_m", aggfunc="first")
        t = t.reindex(columns=[p for p in periods if p in t.columns]).reset_index()
        t["metric"] = metric
        t["_o"] = t.segment.map({s: i for i, s in enumerate(order)}) * 10 + t.channel.map({c: i for i, c in enumerate(channels)})
        out.append(t)
    t = pd.concat(out).sort_values(["metric", "_o"], ascending=[False, True]).drop(columns="_o")
    t.insert(3, "xbrl_tag", t.metric.map({"Revenue": "us-gaap:" + SEGMENT_REVENUE[8:], "EBIT": "nke:EarningsBeforeInterestAndTaxes"}))
    return t[["metric", "segment", "channel", "xbrl_tag"] + [p for p in periods if p in t.columns]]


def wide(sel, stmt, periods):
    """Preview table: one row per line (labels and order from the latest 10-K), periods as columns."""
    s = sel[(sel.statement == stmt) & sel.period.isin(periods)]
    latest_10k = s[s.form == "10-K"].filing_date.max()
    # Latest label/order per line; lines absent from the latest 10-K go to the bottom and are flagged.
    info = s.sort_values(["filing_date", "form"]).drop_duplicates("line_id", keep="last").set_index("line_id")
    in_latest = set(s[s.filing_date == latest_10k].line_id)
    k = s[s.filing_date == latest_10k].drop_duplicates("line_id").set_index("line_id")
    info["rank"] = [k.line_order[i] if i in in_latest else 1000 + info.line_order[i] for i in info.index]
    info["note"] = ["" if i in in_latest else "not in latest 10-K layout" for i in info.index]
    info["label"] = [k.label[i] if i in in_latest else info.label[i] for i in info.index]
    t = s.pivot_table(index="line_id", columns="period", values="value_presented", aggfunc="first")
    t = t.reindex(columns=[p for p in periods if p in t.columns])
    t = info[["label", "rank", "note"]].join(t).sort_values("rank").drop(columns="rank")
    t.insert(1, "xbrl_tag", [i.replace("_", ":", 1) for i in t.index])
    return t.reset_index(drop=True)


# --- Main ---------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--refresh", action="store_true", help="re-download files even if cached")
    args = ap.parse_args()
    os.makedirs(CLEAN_DIR, exist_ok=True)

    filings = list_filings(args.refresh)
    print("Filings used:")
    print(filings[["form", "reportDate", "filingDate", "accessionNumber"]].to_string(index=False))

    all_rows, all_seg, all_sup = [], [], []
    for f in filings.itertuples(index=False):
        paths = download_filing(f, args.refresh)
        rows, facts, labels, reports = extract_statements(f, paths)
        all_rows.append(rows)
        all_seg.append(extract_segments(f, facts, labels, reports, paths["pre"]))
        all_sup.append(extract_supplementary(f, facts))
        print(f"  {f.form} {f.reportDate}: {len(rows)} statement rows, {len(all_seg[-1])} segment facts")

    long = pd.concat(all_rows, ignore_index=True)
    seg = pd.concat(all_seg, ignore_index=True)
    sup = pd.concat(all_sup, ignore_index=True)
    # Keep each item/period from the most recent filing that reports it.
    sup = sup.sort_values("filing_date").drop_duplicates(["concept", "instrument", "period"], keep="last")
    long["value_usd_m"] += 0.0  # turn -0.0 (from negated zero values) into 0.0
    long["value_presented"] += 0.0
    long["label"] = long.label.str.replace(" [Member]", "", regex=False)
    sel = select_latest(long)
    seg_sel = select_segments(seg)
    restated = find_restatements(long)
    checks = pd.concat([run_checks(sel), segment_checks(seg_sel, sel)], ignore_index=True)

    latest_fy = int(filings[filings.form == "10-K"].reportDate.max()[:4])
    fy_flow = [f"FY{y}" for y in range(latest_fy - N_ANNUAL + 1, latest_fy + 1)]
    fy_bs = [f"FY{latest_fy - N_ANNUAL}"] + fy_flow  # one extra opening balance sheet
    q = sel[~sel.period.str.match(r"^FY\d{4}$")].period.unique().tolist()
    q_sorted = sorted(q, key=lambda p: (p.split()[-1], p))

    # utf-8-sig so Excel shows characters such as en dashes in Nike's labels correctly
    save = lambda df, name: df.to_csv(os.path.join(CLEAN_DIR, name), index=False, encoding="utf-8-sig")
    save(filings, "filings.csv")
    save(long, "statements_long.csv")
    save(sel, "statements_selected.csv")
    save(seg, "segments_long.csv")
    save(seg_sel, "segments_selected.csv")
    save(sup, "supplementary.csv")
    save(restated, "restatements.csv")
    save(checks, "checks.csv")
    for stmt, periods in [("IS", fy_flow + q_sorted), ("BS", fy_bs + q_sorted), ("CF", fy_flow + q_sorted)]:
        save(wide(sel, stmt, periods), f"statements_wide_{stmt}.csv")
    save(segments_wide(seg_sel, fy_flow + q_sorted), "segments_wide.csv")

    print(f"\nRows: long={len(long)}, selected={len(sel)}, segment facts={len(seg)} (selected {len(seg_sel)}), "
          f"restatement/layout differences={len(restated)}")
    failed = checks[checks["pass"].eq(False)]
    print(f"Checks: {int(checks['pass'].eq(True).sum())} passed, {len(failed)} failed, "
          f"{int(checks['pass'].isna().sum())} info-only or not applicable")
    if len(failed):
        print(failed.to_string(index=False))
    print(f"Outputs written to {CLEAN_DIR}")
    return 0 if failed.empty else 1


if __name__ == "__main__":
    sys.exit(main())
