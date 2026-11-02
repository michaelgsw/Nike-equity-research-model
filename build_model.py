"""Build nike_model.xlsx (operating model + DCF, comps and valuation summary) from the data in data/clean.

Run data_pipeline.py and market_data.py first. Historical values are written as blue inputs, each row citing its XBRL tag
and each column citing its source filing (row 5). Every forecast cell (FY2027E-FY2031E) is an Excel
formula; the only hardcoded forecast drivers live on the Assumptions sheet.

Year sheets share one layout: A label | B unit | C FY2021A (or a single value) | D-H FY2022A-FY2026A | I-M FY2027E-FY2031E | N notes.
Comps and Summary are free-form tables.
Rows are declared first and laid out in a second pass, so a formula can point at a row on any sheet by key.
"""
import os
import re
from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference, Series
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

HERE = os.path.dirname(os.path.abspath(__file__))
CLEAN = os.path.join(HERE, "data", "clean")
OUT = os.path.join(HERE, "nike_model.xlsx")

YEARS = list(range(2021, 2032))
HIST = list(range(2022, 2027))
BS_HIST = list(range(2021, 2027))
FC = list(range(2027, 2032))
COL = {y: "CDEFGHIJKLM"[i] for i, y in enumerate(YEARS)}
FIRST_ROW = 6

NUM, NUM1, PCT = '#,##0;(#,##0);"-"', '#,##0.0;(#,##0.0);"-"', '0.0%;(0.0%);"-"'
PCT2 = '0.00%;(0.00%);"-"'
DAYS, USD2, MULT = '0.0', '$#,##0.00', '0.0"x"'
BLUE, BLACK, GREEN, GREY = "0000FF", "000000", "008000", "808080"
FC_FILL = PatternFill("solid", fgColor="EEF3FA")
HEAD_FILL = PatternFill("solid", fgColor="1F3864")
SECTION_FILL = PatternFill("solid", fgColor="D9E1F2")
YELLOW = PatternFill("solid", fgColor="FFFF00")
THIN = Side(style="thin", color="000000")


def font(color=BLACK, bold=False, italic=False, size=10):
    return Font(name="Arial", color=color, bold=bold, italic=italic, size=size)


# ---------------------------------------------------------------- data

def load_data():
    st = pd.read_csv(os.path.join(CLEAN, "statements_selected.csv"), encoding="utf-8-sig")
    st = st[st.period.str.match(r"^FY\d{4}$")].copy()
    st["year"] = st.period.str[2:].astype(int)
    seg = pd.read_csv(os.path.join(CLEAN, "segments_wide.csv"), encoding="utf-8-sig")
    seg_sel = pd.read_csv(os.path.join(CLEAN, "segments_selected.csv"), encoding="utf-8-sig")
    sup = pd.read_csv(os.path.join(CLEAN, "supplementary.csv"), encoding="utf-8-sig")
    filings = pd.read_csv(os.path.join(CLEAN, "filings.csv"), encoding="utf-8-sig")
    return st, seg, seg_sel, sup, filings


ST, SEG, SEG_SEL, SUP, FILINGS = load_data()
FILING_LABEL = {r.accessionNumber: f"FY{r.reportDate[2:4]} {r.form}" if r.form == "10-K" else f"Q1 FY{int(r.reportDate[2:4]) + 1} {r.form}"
                for r in FILINGS.itertuples()}
FILING_URL = dict(zip(FILINGS.accessionNumber, FILINGS.document_url))

# Stage 3 inputs (market_data.py): market data, peer financials, monthly returns for beta; Q1 statements for LTM.
MKT = pd.read_csv(os.path.join(CLEAN, "market_inputs.csv"), encoding="utf-8-sig").set_index("item")
PEERS = pd.read_csv(os.path.join(CLEAN, "peers.csv"), encoding="utf-8-sig")
BETA = pd.read_csv(os.path.join(CLEAN, "beta_returns.csv"), encoding="utf-8-sig")
ST_Q = pd.read_csv(os.path.join(CLEAN, "statements_selected.csv"), encoding="utf-8-sig")
ST_Q = ST_Q[ST_Q.period.str.startswith("Q1")]


def mkt(item):
    """Market input and its 'date | source' note."""
    r = MKT.loc[item]
    return float(r.value), f"{r.as_of} | {r.source}"


def q1(statement, concept, period="Q1 FY2027"):
    """Q1 value as presented, with its filing and XBRL tag."""
    r = ST_Q[(ST_Q.statement == statement) & (ST_Q.concept == concept) & (ST_Q.period == period)].iloc[0]
    return float(r.value_presented), f"{r.form} {r.accession}, {period} (ended {r.period_end}), XBRL {concept.replace('_', ':', 1)}"


def stmt(statement, line_id):
    """Historical values (as presented on the face of the statement) for one line, keyed by fiscal year."""
    rows = ST[(ST.statement == statement) & (ST.line_id == line_id)]
    if rows.empty:
        raise KeyError(f"{statement} {line_id} not found in statements_selected.csv")
    tag = line_id.split("[")[0].replace("_", ":", 1)
    return {int(r.year): float(r.value_presented) for r in rows.itertuples()}, f"XBRL {tag}"


def column_sources(statement):
    rows = ST[ST.statement == statement].drop_duplicates("year")
    return dict(zip(rows.year.astype(int), rows.accession))


def seg(metric, segment, channel="All"):
    row = SEG[(SEG.metric == metric) & (SEG.segment == segment) & (SEG.channel == channel)]
    if row.empty:
        raise KeyError(f"segment {metric} {segment} {channel}")
    row = row.iloc[0]
    out = {y: float(row[f"FY{y}"]) for y in HIST if pd.notna(row[f"FY{y}"])}
    return out, f"XBRL {row.xbrl_tag}" if isinstance(row.xbrl_tag, str) else "sum of channels"


def seg_q1(segment, channel="All"):
    row = SEG[(SEG.metric == "Revenue") & (SEG.segment == segment) & (SEG.channel == channel)].iloc[0]
    return float(row["Q1 FY2026"]), float(row["Q1 FY2027"])


def sup(item, instrument_contains=None):
    rows = SUP[SUP.item == item]
    if instrument_contains:
        rows = rows[rows.instrument.fillna("").str.contains(instrument_contains)]
    rows = rows[rows.period.str.match(r"^FY\d{4}$")]
    return {int(p[2:]): float(v) for p, v in zip(rows.period, rows.value)}, rows


# ---------------------------------------------------------------- row registry

REG = {}          # (sheet, key) -> row number
CUR = {"sheet": None}


def R(key, y, col_abs=False):
    """Cell address of row `key` in fiscal year y; 'Sheet.key' points at another sheet."""
    sheet, k = key.split(".", 1) if "." in key else (CUR["sheet"], key)
    c = ("$" if col_abs else "") + COL[y]
    addr = f"{c}{REG[(sheet, k)]}"
    return addr if sheet == CUR["sheet"] else f"{sheet}!{addr}"


def S(key, sheet="Assumptions"):
    """Absolute address of a scalar (column C), on the Assumptions sheet unless another sheet is named."""
    addr = f"$C${REG[(sheet, key)]}"
    return addr if CUR["sheet"] == sheet else f"{sheet}!{addr}"


def DEBT(col):
    """Absolute range of one column of the bond table on the Assumptions sheet."""
    a, b = REG[("Assumptions", "debt_first")], REG[("Assumptions", "debt_last")]
    rng = f"${col}${a}:${col}${b}"
    return rng if CUR["sheet"] == "Assumptions" else f"Assumptions!{rng}"


class Row:
    def __init__(self, key, label, unit="$m", hist=None, fc=None, fmt=NUM, note="", bold=False,
                 indent=0, total=False, hist_years=None):
        self.key, self.label, self.unit, self.hist, self.fc = key, label, unit, hist, fc
        self.fmt, self.note, self.bold, self.indent, self.total = fmt, note, bold or total, indent, total
        self.hist_years = hist_years


class Section:
    def __init__(self, text):
        self.text = text


class Text:
    def __init__(self, text, italic=False):
        self.text, self.italic = text, italic


class Scalar:
    def __init__(self, key, label, unit, value, fmt=NUM, note="", key_input=False):
        self.key, self.label, self.unit, self.value, self.fmt, self.note, self.key_input = key, label, unit, value, fmt, note, key_input


class Block:
    """Free-form table: render(ws, first_row) writes `nrows` rows; `keys` maps names to row offsets."""
    def __init__(self, nrows, render, keys=None):
        self.nrows, self.render, self.keys = nrows, render, keys or {}


class Sheet:
    def __init__(self, name, title, hist_years=HIST, sources=None, show_2021=False, c_label=None,
                 years=True, subtitle=None, widths=None, freeze="C6"):
        self.name, self.title, self.hist_years, self.sources, self.show_2021 = name, title, hist_years, sources or {}, show_2021
        self.c_label = c_label    # header for column C when it holds single values instead of FY2021
        self.years = years        # False for free-form sheets (Comps, Summary): no year header, not audited
        self.subtitle, self.widths, self.freeze = subtitle, widths or {}, freeze
        self.items = []

    def row(self, *a, **k):
        self.items.append(Row(*a, **k))

    def section(self, text):
        self.items.append(Section(text))

    def blank(self):
        self.items.append(None)

    def layout(self):
        r = FIRST_ROW
        for it in self.items:
            if isinstance(it, (Row, Scalar)):
                REG[(self.name, it.key)] = r
            if isinstance(it, Block):
                for k, off in it.keys.items():
                    REG[(self.name, k)] = r + off
                r += it.nrows
            else:
                r += 1


# ---------------------------------------------------------------- sheets

def growth_fc(key, driver):
    return lambda y: f"={R(key, y - 1)}*(1+{R(driver, y)})"


def growth_rate(key, first=2023):
    return lambda y: None if y < first else f"=IF({R(key, y - 1)}=0,0,{R(key, y)}/{R(key, y - 1)}-1)"


def ratio(num, den):
    return lambda y: f"=IF({R(den, y)}=0,0,{R(num, y)}/{R(den, y)})"


def link(key):
    return lambda y: f"={R(key, y)}"


def prior(key):
    return lambda y: f"={R(key, y - 1)}"


def total(keys):
    return lambda y: "=" + "+".join(R(k, y) for k in keys)


def subtract(a, b):
    return lambda y: f"={R(a, y)}-{R(b, y)}"


GEOS = [("na", "North America"), ("emea", "EMEA"), ("gc", "Greater China"), ("apla", "APLA")]
CHANNELS = [("w", "Wholesale"), ("d", "NIKE Direct")]

# Scenario drivers, FY2027E-FY2031E, in percent. Analyst assumptions; rationale in the notes column.
GROWTH = {
    "na_w": ([5, 4, 4, 3.5, 3.5], [7, 5, 5, 4, 4], [2, 2, 2, 2, 2]),
    "na_d": ([-4, 2, 4, 5, 5], [0, 6, 7, 6, 5], [-8, -2, 0, 1, 2]),
    "emea_w": ([0, 3, 3, 3, 3], [2, 4, 4, 4, 4], [-3, 1, 1, 2, 2]),
    "emea_d": ([-8, 2, 4, 4, 4], [-4, 6, 6, 5, 5], [-12, -3, 0, 1, 2]),
    "gc_w": ([-20, 0, 4, 5, 5], [-13, 6, 8, 7, 6], [-27, -8, -3, 0, 2]),
    "gc_d": ([-12, 0, 4, 5, 5], [-7, 7, 9, 8, 6], [-17, -6, -2, 0, 2]),
    "apla_w": ([1, 4, 5, 5, 5], [3, 6, 6, 6, 5], [-2, 2, 2, 3, 3]),
    "apla_d": ([-2, 3, 5, 6, 6], [1, 7, 7, 7, 6], [-5, 0, 1, 2, 3]),
    "conv": ([-22, -5, 0, 2, 2], [-16, 3, 5, 5, 4], [-29, -12, -5, 0, 0]),
}
GROWTH_NOTES = {
    "na_w": "Q1 FY27 +9.0% YoY as wholesale partners are reinvested; normalizes to low-single digits.",
    "na_d": "Q1 FY27 -6.0%: digital repositioned to full price (fewer promotions); returns to growth from FY28 in Base.",
    "emea_w": "Q1 FY27 -1.2% YoY; flat FY27 then GDP-plus growth.",
    "emea_d": "Q1 FY27 -11.9% YoY; FY27E set partway to the Q1 run-rate (H2 assumed better as the reset ends); same full-price repositioning as North America.",
    "gc_w": "Q1 FY27 -27.9% YoY; FY27E set partway to the Q1 run-rate (H2 assumed better as the reset ends); 10-K expects Greater China weakness through FY2027.",
    "gc_d": "Q1 FY27 -13.4% YoY; FY27E set partway to the Q1 run-rate (H2 assumed better as the reset ends); stabilizes FY28 in Base, keeps shrinking in Bear.",
    "apla_w": "Q1 FY27 -0.3% YoY.",
    "apla_d": "Q1 FY27 -4.4% YoY.",
    "conv": "Q1 FY27 -28.1% YoY; FY27E set partway to the Q1 run-rate (H2 assumed better as the reset ends); 10-K expects Converse weakness through FY2027.",
}
UNDERLYING_GM = ([43.3, 43.6, 44.2, 44.6, 44.8], [43.7, 44.4, 45.2, 45.8, 46.0], [42.7, 42.6, 42.8, 43.0, 43.0])
TARIFF = ([-0.5] * 5, [0.0] * 5, [-1.5, -1.2, -1.0, -1.0, -1.0])
NARRATIVE = {
    "Base": "Marketplace reset finishes by Dec-2026; NIKE Direct returns to growth from FY2028 as the full-price digital "
            "strategy matures, while Greater China and Converse fall again in FY2027 before stabilizing. Gross margin rebuilds "
            "slowly as reset costs roll off, less a 50 bp allowance for non-IEEPA / replacement tariffs.",
    "Bull": "Wholesale reinvestment and product cadence drive a faster recovery: NIKE Direct is flat in FY2027 and grows "
            "from FY2028, Greater China grows again from FY2028, no net tariff drag, and gross margin moves back toward "
            "the FY2022-FY2024 range (43.5%-46.0%).",
    "Bear": "Greater China and Converse keep shrinking through FY2029 and NIKE Direct declines further as promotions are "
            "pulled without a wholesale offset; new tariffs cost 100-150 bp of gross margin and the reset adds little margin.",
}


def per_year(values, pct=True):
    return {y: (v / 100 if pct else v) for y, v in zip(FC, values)}


def build_assumptions():
    sh = Sheet("Assumptions", "NIKE, Inc. - Assumptions and scenario drivers", c_label="Input")
    it = sh.items
    sh.section("Scenario and market data")
    it.append(Scalar("scenario", "Scenario selector (Base / Bull / Bear)", "text", "Base", fmt="@",
                     note="Pick from the dropdown; every scenario row below switches on this cell.", key_input=True))
    it.append(Scalar("scen_idx", "Scenario index (1 = Base, 2 = Bull, 3 = Bear)", "#",
                     lambda: f'=IF({S("scenario")}="Bull",2,IF({S("scenario")}="Bear",3,1))', fmt="0"))
    price, price_src = mkt("price_NKE")
    it.append(Scalar("price", "NKE share price (close)", "$/share", price, fmt=USD2,
                     note=f"Close {price_src}. Nasdaq historical quotes give the same close (${mkt('price_NKE_nasdaq_check')[0]:.2f})."))
    it.append(Scalar("price_date", "Pricing and valuation date", "date", datetime(2026, 10, 30), fmt="yyyy-mm-dd",
                     note="Last close before the 2026-11-02 build date. The DCF discounts cash flows to this date."))
    cls_a = SUP[(SUP.item == "Shares outstanding (cover page)") & (SUP.period == "2026-09-28") & (SUP.value < 1000)].value.iloc[0]
    cls_b = SUP[(SUP.item == "Shares outstanding (cover page)") & (SUP.period == "2026-09-28") & (SUP.value > 1000)].value.iloc[0]
    it.append(Scalar("shares_a", "Class A shares outstanding at 2026-09-28", "m", round(cls_a, 3), fmt=NUM1,
                     note="Q1 FY27 10-Q 0000320187-26-000193 cover page, XBRL dei:EntityCommonStockSharesOutstanding (Class A)."))
    it.append(Scalar("shares_b", "Class B shares outstanding at 2026-09-28", "m", round(cls_b, 3), fmt=NUM1,
                     note="Q1 FY27 10-Q 0000320187-26-000193 cover page, XBRL dei:EntityCommonStockSharesOutstanding (Class B)."))
    it.append(Scalar("shares_out", "Total shares outstanding at 2026-09-28", "m",
                     lambda: f"={S('shares_a')}+{S('shares_b')}", fmt=NUM1))
    for key, label, unit, item, fmt in (("low_52w", "NKE 52-week low (daily close)", "$/share", "nke_52w_low", USD2),
                                        ("high_52w", "NKE 52-week high (daily close)", "$/share", "nke_52w_high", USD2),
                                        ("fx_usd_eur", "EUR/USD reference rate", "USD/EUR", "fx_usd_per_eur", "0.0000"),
                                        ("fx_chf_eur", "EUR/CHF reference rate", "CHF/EUR", "fx_chf_per_eur", "0.0000")):
        v, src = mkt(item)
        it.append(Scalar(key, label, unit, v, fmt=fmt, note=src))
    sh.blank()

    sh.section("Financing and other scalar assumptions")
    it.append(Scalar("direct_premium", "NIKE Direct gross margin premium vs. wholesale", "pts", 0.10, fmt=PCT,
                     note="Analyst assumption: Nike does not disclose margin by channel. Drives the channel-mix effect on gross margin."))
    it.append(Scalar("new_debt_rate", "Coupon on new / refinancing debt", "%", lambda: f"={S('kd_pre')}", fmt=PCT2,
                     note="Linked to the pre-tax cost of debt in the valuation inputs below (Treasury + A2/A spread). Stage 2 used a 4.75% placeholder."))
    it.append(Scalar("refi_pct", "Share of maturing notes refinanced", "%", 1.0, fmt=PCT,
                     note="Assumes Nike refinances maturities (it has kept ~$8bn of notes outstanding for years). New notes assumed issued mid-year."))
    it.append(Scalar("cash_yield", "Interest yield on beginning cash + short-term investments", "%", 0.0325, fmt=PCT,
                     note="Close to FY2026 implied yield (see row below). Uses beginning-of-year balances so the model has no circular reference."))
    it.append(Scalar("implied_yield", "FY2026 implied yield (memo)", "%",
                     lambda: f"=({R('BS.int_exist', 2026)}-{R('IS.int', 2026)})/({R('BS.cash', 2025)}+{R('BS.sti', 2025)})", fmt=PCT,
                     note="(Coupon interest on notes - reported net interest expense) / beginning cash + STI. Approximate: ignores fees, leases and capitalized interest."))
    sh.blank()

    sh.section("Valuation inputs (WACC, terminal value, sensitivity steps; one set for all scenarios)")
    rf, rf_src = mkt("rf_10y")
    it.append(Scalar("rf", "Risk-free rate: 10-year US Treasury par yield", "%", rf, fmt=PCT2, note=rf_src))
    erp, erp_src = mkt("erp_implied")
    it.append(Scalar("erp", "Equity risk premium (implied, S&P 500)", "%", erp, fmt=PCT2, note=erp_src))
    it.append(Scalar("blume", "Blume adjustment: weight on raw beta", "x", 0.67, fmt="0.00",
                     note="Adjusted beta = 0.67 x raw + 0.33 x 1.0 (Blume 1971; the convention Bloomberg uses). Raw beta is regressed on the DCF sheet."))
    spread, spread_src = mkt("spread_A2_A")
    it.append(Scalar("spread", "Default spread: Moody's A2 / S&P A", "%", spread, fmt=PCT2,
                     note=f"Nike is rated A+ (S&P) / A2 (Moody's) per the FY26 10-K (0000320187-26-000088), liquidity section. Split rating: "
                          f"uses the lower (A2) bucket; A1/A+ would be {mkt('spread_A1_Aplus')[0]:.2%}. {spread_src}"))
    it.append(Scalar("kd_pre", "Pre-tax cost of debt = risk-free rate + default spread", "%", lambda: f"={S('rf')}+{S('spread')}", fmt=PCT2))
    it.append(Scalar("tg", "Terminal growth rate (perpetuity method)", "%", 0.0225, fmt=PCT2, key_input=True,
                     note="Analyst assumption: roughly inflation-level growth for a mature global brand; below the 10-year Treasury yield and "
                          "long-run nominal GDP growth. Lowered from 2.5% so the perpetuity and exit-multiple values sit closer "
                          "(see the terminal value cross-check on the DCF sheet)."))
    it.append(Scalar("exit_mult", "Exit multiple: EV / FY2031E EBITDA", "x", 10.0, fmt=MULT, key_input=True,
                     note="Analyst assumption: above the peer LTM median (Comps sheet) for Nike's scale and brand, below Nike's own LTM multiple, which is inflated by trough earnings."))
    it.append(Scalar("wacc_step", "Sensitivity step: WACC", "%", 0.005, fmt=PCT2))
    it.append(Scalar("g_step", "Sensitivity step: terminal growth", "%", 0.0025, fmt=PCT2))
    it.append(Scalar("mult_step", "Sensitivity step: exit multiple", "x", 1.0, fmt=MULT))
    it.append(Scalar("margin_step", "Sensitivity step: FY2030E EBIT margin", "%", 0.01, fmt=PCT2))
    sh.blank()

    sh.section("Scenario narratives")
    for name, txt in NARRATIVE.items():
        it.append(Text(f"{name}: {txt}"))
    sh.blank()

    def scenario(key, label, unit, cases, hist=None, note="", fmt=PCT):
        for name, vals in zip(("Base", "Bull", "Bear"), cases):
            sh.row(f"{key}_{name.lower()}", f"{label} - {name}", unit, fc=per_year(vals) if fmt == PCT else per_year(vals, False),
                   fmt=fmt, indent=1)
        sh.row(key, f"{label} - LIVE", unit, hist=hist,
               fc=lambda y: f"=CHOOSE({S('scen_idx')},{R(key + '_base', y)},{R(key + '_bull', y)},{R(key + '_bear', y)})",
               fmt=fmt, bold=True, note=note)

    sh.section("Revenue growth by geography and channel (YoY, constant mix of reporting)")
    for g, gname in GEOS:
        for c, cname in CHANNELS:
            k = f"{g}_{c}"
            scenario(f"g_{k}", f"{gname} {cname} growth", "%", GROWTH[k], hist=link(f"Segments.{k}_g"), note=GROWTH_NOTES[k])
    scenario("g_conv", "Converse revenue growth", "%", GROWTH["conv"], hist=link("Segments.conv_g"), note=GROWTH_NOTES["conv"])
    sh.row("corp_rev", "Corporate revenue (FX hedge gains/losses)", "$m", hist=link("Segments.corp"), fc=per_year([0] * 5, False),
           note="Not forecastable (hedge results); set to zero. Global Brand Divisions revenue is held flat on the Segments sheet.")
    sh.blank()

    sh.section("Gross margin build: underlying + tariffs + channel mix")
    scenario("gm_under", "Underlying gross margin (tariff-neutral)", "%", UNDERLYING_GM, hist=link("IS.gm"),
             note="FY2026 42.9% is roughly tariff-neutral: the $986m IEEPA refund 'largely offset' IEEPA tariffs (FY26 10-K MD&A). "
                  "Q1 FY27 GM 42.8% (+60 bp); Base FY27 set so modeled GM is in line with Q1. Rebuild reflects reset costs rolling off and fewer promotions.")
    scenario("gm_tariff", "Net tariff impact vs. FY2026 baseline", "%", TARIFF,
             note="Analyst assumption: allowance for non-IEEPA tariffs or replacement tariffs after the Feb-20-2026 Supreme Court ruling.")
    sh.row("gm_mix", "Channel-mix effect", "%",
           fc=lambda y: f"=({R('Segments.direct_share', y)}-{R('Segments.direct_share', 2026, col_abs=True)})*{S('direct_premium')}",
           note="(NIKE Direct share of NIKE Brand revenue - FY2026 share) x Direct margin premium.")
    sh.row("gm", "Gross margin - modeled", "%", hist=link("IS.gm"),
           fc=lambda y: f"={R('gm_under', y)}+{R('gm_tariff', y)}+{R('gm_mix', y)}", fmt=PCT, bold=True)
    sh.blank()

    sh.section("Operating expenses (single path)")
    sh.row("dc_pct", "Demand creation, % of revenue", "%", hist=ratio("IS.dc", "IS.rev"),
           fc=per_year([10.2, 10.0, 9.8, 9.6, 9.5]), fmt=PCT,
           note="FY2026 10.2%; drifts down as Pace savings and wholesale (partner-funded marketing) mix help.")
    sh.row("oh_pct", "Operating overhead ex-restructuring, % of revenue", "%",
           hist=lambda y: f"=({R('IS.oh', y)}-{R('IS.restr_oh', y)})/{R('IS.rev', y)}",
           fc=per_year([24.0, 23.5, 23.0, 22.6, 22.3]), fmt=PCT,
           note="FY2026 24.0% ex-severance. Pace program (Oct-2026): ~$2.5bn cumulative savings through FY2031 before reinvestment (Q1 FY27 10-Q).")
    sh.row("restr", "Pace restructuring charges (in operating overhead)", "$m", hist=link("IS.restr_oh"),
           fc=per_year([300, 200, 200, 150, 150], False),
           note="~$1.0bn total pre-tax, ~$0.3bn in FY2027 'largely operating overhead' (Q1 FY27 10-Q). FY28-31 phasing is an analyst assumption.")
    sh.row("oth", "Other (income) expense, net", "$m", hist=link("IS.oth"), fc=per_year([0] * 5, False),
           note="Mostly FX and hedge gains/losses; not forecast.")
    sh.row("tax", "Effective tax rate", "%", hist=link("IS.etr"), fc=per_year([20.0] * 5), fmt=PCT,
           note="FY2026 ETR 20.3%; Q1 FY27 22.7%. FY2022-25 ranged 9.1%-18.2% with discrete benefits.")
    sh.blank()

    sh.section("Working capital (days on ending balances, 365-day year)")
    it.append(Scalar("ieepa", "IEEPA tariff refund receivable in FY2026 accounts receivable", "$m", 684,
                     note="FY26 10-K (0000320187-26-000088), Note 1 and MD&A tariffs discussion: $684m of outstanding IEEPA tariff "
                          "receivables within Accounts receivable, net at May 31, 2026; 'substantially all' received after May 31, 2026 "
                          "(Q1 FY27 AR fell $695m). Excluded from the DSO starting point and from DCF cash flows (already in 2026-08-31 cash)."))
    it.append(Scalar("dio_target", "Target inventory days (DIO)", "days", 98.0, fmt=DAYS, key_input=True,
                     note="Analyst assumption. History: FY2022 121.8, FY2023 106.7, FY2024 96.4 (post-destocking low), FY2025 103.1, "
                          "FY2026 103.4. Target sits between the FY2024 low and today's level."))
    it.append(Scalar("dso_target", "Target receivable days (DSO)", "days", 37.0, fmt=DAYS,
                     note="Analyst assumption: FY2025 level (37.2). FY2022-24 29-37 days."))
    it.append(Scalar("wc_end", "Fiscal year the working capital targets are reached", "FY", 2028, fmt="0",
                     note="FY2027E closes half of the gap from the FY2026 starting point; FY2028E onward sits at target. "
                          "Spreads the working capital release over FY2027-28 instead of front-loading it."))
    it.append(Scalar("dso_start", "Starting DSO: FY2026 excluding the IEEPA receivable", "days",
                     lambda: f"=({R('BS.ar', 2026)}-{S('ieepa')})/{R('IS.rev', 2026)}*365", fmt=DAYS))
    it.append(Scalar("dio_start", "Starting DIO: FY2026 actual", "days", lambda: f"={R('Ratios.dio', 2026)}", fmt=DAYS))
    sh.row("wc_glide", "Share of the gap to target still open", "%", fmt=PCT,
           fc=lambda y: f"=MAX(0,{S('wc_end')}-{R('BS.fy', y)})/({S('wc_end')}-{R('BS.fy', 2027, True)}+1)",
           note="Share = (target year - fiscal year) / (target year - FY2027 + 1), floored at 0: 50% in FY2027E, 0% from FY2028E.")
    sh.row("dso", "Days sales outstanding (AR / revenue)", "days", hist=link("Ratios.dso"), fmt=DAYS,
           fc=lambda y: f"={S('dso_target')}+({S('dso_start')}-{S('dso_target')})*{R('wc_glide', y)}",
           note="Glides from the FY2026 level excluding the $684m IEEPA receivable (FY2026 reported 46.7 days) to the target.")
    sh.row("dio", "Days inventory outstanding (inventory / COGS)", "days", hist=link("Ratios.dio"), fmt=DAYS,
           fc=lambda y: f"={S('dio_target')}+({S('dio_start')}-{S('dio_target')})*{R('wc_glide', y)}",
           note="Glides from FY2026 actual to the target inventory days above.")
    sh.row("dpo", "Days payables outstanding (AP / COGS)", "days", hist=link("Ratios.dpo"), fc=per_year([48] * 5, False), fmt=DAYS,
           note="FY2025-26 48-50 days; AP includes ~$1.1bn supplier-finance obligations (FY26 10-K).")
    sh.row("prepaid_pct", "Prepaid expenses and other current assets, % of revenue", "%", hist=ratio("BS.prepaid", "IS.rev"),
           fc=per_year([4.5] * 5), fmt=PCT, note="FY2022-26 range 3.6%-4.6%.")
    sh.row("accrued_pct", "Accrued liabilities, % of revenue", "%", hist=ratio("BS.accrued", "IS.rev"),
           fc=per_year([12.8] * 5), fmt=PCT, note="FY2022-26 range 11.1%-13.3%.")
    sh.blank()

    sh.section("Capex, D&A and stock-based compensation")
    sh.row("capex_pct", "Capital expenditures, % of revenue", "%", hist=lambda y: f"=-{R('CF.capex', y)}/{R('IS.rev', y)}",
           fc=per_year([1.6] * 5), fmt=PCT, note="FY2022-26 range 0.9%-1.9%; average ~1.5%.")
    sh.row("da_pct", "Depreciation and amortization, % of revenue", "%", hist=ratio("IS.da", "IS.rev"),
           fc=per_year([1.6] * 5), fmt=PCT, note="FY2023-26 1.4%-1.7%. FY2022 shows depreciation only (old cash-flow layout).")
    sh.row("sbc_pct", "Stock-based compensation, % of revenue", "%", hist=ratio("CF.sbc", "IS.rev"),
           fc=per_year([1.5] * 5), fmt=PCT, note="FY2022-26 range 1.4%-1.6%.")
    sh.blank()

    sh.section("Capital returns and share count")
    sh.row("dps_g", "Dividend per share growth", "%", hist=growth_rate("IS.dps", 2023), fc=per_year([2.5] * 5), fmt=PCT,
           note="FY2026 DPS growth 3.8%; Q1 FY27 quarterly dividend $0.41 vs $0.40 (+2.5%).")
    sh.row("buyback", "Share repurchases", "$m", hist=lambda y: f"=-{R('CF.buyback', y)}",
           fc=per_year([0, 2000, 3000, 4000, 4000], False),
           note="Buybacks paused since Q1 FY26 under the $18bn program; assumed to restart in FY2028 and step up toward the FY22-24 pace "
                "($4-5bn a year), keeping cash near $8-10bn.")
    # EPS is measured on the prior-year share count: using the same year's shares would be circular
    # (shares bought depend on the price, and the price would depend on the shares).
    sh.row("bb_price", "Average repurchase price (grows with EPS)", "$/share", fmt=USD2,
           fc=lambda y: f"={S('price')}*({R('IS.ni', y)}/{R('IS.diluted', y - 1)})/({R('IS.ni', 2027, True)}/{R('IS.diluted', 2026, True)})",
           note="Current price x EPS growth since FY2027, i.e. a constant P/E (equals the current price in FY2027). "
                "EPS here = net income / prior-year diluted shares, which avoids a circular reference.")
    sh.row("share_iss", "Shares issued to employees, % of prior-year shares", "%",
           fc=per_year([0.5] * 5), fmt=PCT, note="Analyst assumption for option exercises and RSU vesting (not separately extracted).")
    sh.row("options", "Proceeds from option exercises and other issuance", "$m", hist=link("CF.options"),
           fc=per_year([350] * 5, False), note="FY2026 $354m.")
    sh.blank()

    def debt_render(ws, r0):
        hdr = ["Senior note", "", "Face $m", "Coupon", "Maturity", "Maturity FY", "Carrying $m"]
        for i, h in enumerate(hdr):
            c = ws.cell(r0, 1 + i, h)
            c.font, c.fill = font(bold=True), SECTION_FILL
        ws.cell(r0, 14, "Source").font = font(bold=True)
        ws.cell(r0, 14).fill = SECTION_FILL
        for i, b in enumerate(BONDS):
            r = r0 + 1 + i
            ws.cell(r, 1, b["name"]).font = font()
            for col, val, fmt in ((3, b["face"], NUM), (4, b["coupon"], '0.000%'), (5, b["maturity"], "@"),
                                  (6, b["fy"], "0"), (7, b["carrying"], NUM)):
                c = ws.cell(r, col, val)
                c.font, c.number_format = font(BLUE), fmt
            ws.cell(r, 14, b["source"]).font = font(GREY, size=8)
        r = r0 + 1 + len(BONDS)
        ws.cell(r, 1, "Total / weighted average coupon").font = font(bold=True)
        for col, f, fmt in ((3, f"=SUM(C{r0 + 1}:C{r - 1})", NUM), (4, f"=SUMPRODUCT(C{r0 + 1}:C{r - 1},D{r0 + 1}:D{r - 1})/C{r}", '0.000%'),
                            (7, f"=SUM(G{r0 + 1}:G{r - 1})", NUM)):
            c = ws.cell(r, col, f)
            c.font, c.number_format, c.border = font(bold=True), fmt, Border(top=THIN)
    sh.section("Senior notes outstanding at May 31, 2026 (FY26 10-K debt note)")
    it.append(Block(len(BONDS) + 2, debt_render, {"debt_first": 1, "debt_last": len(BONDS)}))
    return sh


def parse_bonds():
    """Bond table at May 31, 2026 from the FY26 10-K debt note (face, coupon, carrying), maturity from the instrument name."""
    months = {m: i for i, m in enumerate(["January", "February", "March", "April", "May", "June", "July", "August",
                                          "September", "October", "November", "December"], 1)}
    d = SUP[(SUP.period == "FY2026") & SUP.item.str.startswith("Debt ") & SUP.instrument.notna()]
    out = []
    for inst, g in d.groupby("instrument"):
        m = re.search(r"Due([A-Za-z]+?)(\d{1,2})(\d{4})Member", inst)
        month, day, year = months[m.group(1)], int(m.group(2)), int(m.group(3))
        vals = dict(zip(g.item, g.value))
        out.append({"name": f"{vals['Debt coupon rate']:.2%} notes due {m.group(1)[:3]} {day}, {year}",
                    "face": vals["Debt face amount"], "coupon": vals["Debt coupon rate"], "carrying": vals["Debt carrying amount"],
                    "maturity": f"{year}-{month:02d}-{day:02d}", "fy": year if month <= 5 else year + 1,
                    "source": f"FY26 10-K {g.accession.iloc[0]} | XBRL DebtInstrumentFaceAmount / InterestRateStatedPercentage / LongTermDebt [{inst}]"})
    return sorted(out, key=lambda b: b["maturity"])


BONDS = parse_bonds()


def build_segments():
    accs = SEG_SEL[SEG_SEL.period.str.match(r"^FY\d{4}$")].drop_duplicates("period")
    sh = Sheet("Segments", "NIKE, Inc. - Revenue by geography and channel; segment EBIT",
               sources={int(p[2:]): a for p, a in zip(accs.period, accs.accession)})
    sh.section("NIKE Brand revenue by geography and channel")
    for g, gname in GEOS:
        for c, cname in CHANNELS:
            h, note = seg("Revenue", gname, cname)
            sh.row(f"{g}_{c}", f"{gname} - {cname}", hist=h, fc=growth_fc(f"{g}_{c}", f"Assumptions.g_{g}_{c}"), indent=1, note=note)
        sh.row(g, f"{gname} total", hist=total([f"{g}_w", f"{g}_d"]), fc=total([f"{g}_w", f"{g}_d"]), total=True)
    h, note = seg("Revenue", "Global Brand Divisions")
    sh.row("gbd", "Global Brand Divisions", hist=h, fc=prior("gbd"), note=note + " (licensing and other; held flat)")
    for c, cname in CHANNELS:
        f = total([f"{g}_{c}" for g, _ in GEOS])
        sh.row(f"nb_{c}", f"NIKE Brand {cname}", hist=f, fc=f, indent=1)
    sh.row("nb", "NIKE Brand revenue", hist=total(["nb_w", "nb_d", "gbd"]), fc=total(["nb_w", "nb_d", "gbd"]), total=True)
    sh.row("conv", "Converse revenue", hist=total(["conv_w", "conv_d", "conv_o"]), fc=growth_fc("conv", "Assumptions.g_conv"),
           total=True, note="Forecast as one line; channel split shown for history only.")
    for c, cname, ch in (("w", "Wholesale", "Wholesale"), ("d", "Direct to consumer", "NIKE Direct"), ("o", "Other (licensing)", "Other")):
        h, note = seg("Revenue", "Converse", ch)
        sh.row(f"conv_{c}", f"Converse - {cname}", hist=h, indent=1, note=note + " (memo, history only)")
    h, note = seg("Revenue", "Corporate")
    sh.row("corp", "Corporate (FX hedge results)", hist=h, fc=link("Assumptions.corp_rev"), note=note)
    sh.row("total", "Total NIKE, Inc. revenue", hist=total(["nb", "conv", "corp"]), fc=total(["nb", "conv", "corp"]), total=True,
           note="Ties to IS revenue (see Checks).")
    sh.row("direct_share", "NIKE Direct share of NIKE Brand revenue", "%", hist=ratio("nb_d", "nb"), fc=ratio("nb_d", "nb"), fmt=PCT)
    sh.blank()

    sh.section("Revenue growth (YoY)")
    for g, gname in GEOS:
        for c, cname in CHANNELS:
            sh.row(f"{g}_{c}_g", f"{gname} - {cname}", "%", hist=growth_rate(f"{g}_{c}"), fc=growth_rate(f"{g}_{c}"), fmt=PCT, indent=1)
        sh.row(f"{g}_g", f"{gname} total", "%", hist=growth_rate(g), fc=growth_rate(g), fmt=PCT)
    for k, label in (("nb_d", "NIKE Brand Direct"), ("nb_w", "NIKE Brand Wholesale"), ("nb", "NIKE Brand"), ("conv", "Converse"), ("total", "Total NIKE, Inc.")):
        sh.row(f"{k}_g", label, "%", hist=growth_rate(k), fc=growth_rate(k), fmt=PCT, bold=k == "total")
    sh.blank()

    sh.section("Segment EBIT (history as reported; EBIT = pre-tax income + net interest expense)")
    for g, gname in GEOS + [("gbd", "Global Brand Divisions")]:
        h, note = seg("EBIT", gname)
        sh.row(f"{g}_ebit", gname, hist=h, note=note, indent=1)
    sh.row("nb_ebit", "NIKE Brand EBIT", hist=total([f"{g}_ebit" for g, _ in GEOS] + ["gbd_ebit"]), total=True,
           note="Sum of geographies + GBD (the FY2022 NIKE Brand subtotal is not XBRL-tagged; FY23-26 match the reported subtotal).")
    for k, name in (("conv", "Converse"), ("corp", "Corporate")):
        h, note = seg("EBIT", name)
        sh.row(f"{k}_ebit", name, hist=h, note=note, indent=1)
    sh.row("tot_ebit", "Total NIKE, Inc. EBIT", hist=total(["nb_ebit", "conv_ebit", "corp_ebit"]), total=True,
           note="Ties to IS EBIT (see Checks).")
    for g, gname in GEOS:
        sh.row(f"{g}_ebit_m", f"{gname} EBIT margin", "%", hist=ratio(f"{g}_ebit", g), fmt=PCT, indent=1)
    sh.row("conv_ebit_m", "Converse EBIT margin", "%", hist=ratio("conv_ebit", "conv"), fmt=PCT, indent=1)
    sh.blank()

    q1_lines = [(f"{g}_{c}", f"{gname} - {cname}", gname, cname) for g, gname in GEOS for c, cname in CHANNELS]
    q1_lines.append(("conv", "Converse", "Converse", "All"))

    def q1_render(ws, r0):
        for i, h in enumerate(["Q1 sanity check: 3 months ended Aug 31 (Q1 FY27 10-Q)", "", "", "Q1 FY2026A", "Q1 FY2027A",
                               "Q1 YoY", "FY2027E growth (live)", "Gap, pts"]):
            c = ws.cell(r0, 1 + i, h)
            c.font, c.fill = font(bold=True), SECTION_FILL
        ws.cell(r0, 14, "Q1 YoY vs. full-year FY2027E assumption; a large gap flags a driver to revisit.").font = font(GREY, size=8)
        CUR["sheet"] = "Segments"
        for i, (key, label, segname, ch) in enumerate(q1_lines):
            r = r0 + 1 + i
            q26, q27 = seg_q1(segname, ch)
            ws.cell(r, 1, label).font = font()
            for col, val in ((4, q26), (5, q27)):
                c = ws.cell(r, col, val)
                c.font, c.number_format = font(BLUE), NUM
            driver = "g_conv" if key == "conv" else f"g_{key}"
            for col, f, color in ((6, f"=E{r}/D{r}-1", BLACK), (7, f"={R('Assumptions.' + driver, 2027)}", GREEN), (8, f"=G{r}-F{r}", BLACK)):
                c = ws.cell(r, col, f)
                c.font, c.number_format = font(color), PCT
            ws.cell(r, 14, "Q1 FY27 10-Q 0000320187-26-000193, segment revenue note").font = font(GREY, size=8)
    sh.items.append(Block(len(q1_lines) + 1, q1_render))
    return sh


def build_is():
    sh = Sheet("IS", "NIKE, Inc. - Income statement", sources=column_sources("IS"))
    IS = lambda lid: stmt("IS", lid)
    h, n = IS("us-gaap_RevenueFromContractWithCustomerExcludingAssessedTax")
    sh.row("rev", "Revenues", hist=h, fc=link("Segments.total"), note=n + "; forecast = Segments build", total=True)
    sh.row("rev_g", "Revenue growth", "%", hist=growth_rate("rev"), fc=growth_rate("rev"), fmt=PCT, indent=1)
    h, n = IS("us-gaap_CostOfGoodsAndServicesSold")
    sh.row("cogs", "Cost of sales", hist=h, fc=lambda y: f"={R('rev', y)}*(1-{R('Assumptions.gm', y)})", note=n)
    sh.row("gp", "Gross profit", hist=subtract("rev", "cogs"), fc=subtract("rev", "cogs"), total=True)
    sh.row("gm", "Gross margin", "%", hist=ratio("gp", "rev"), fc=ratio("gp", "rev"), fmt=PCT, indent=1)
    h, n = IS("us-gaap_MarketingAndAdvertisingExpense")
    sh.row("dc", "Demand creation expense", hist=h, fc=lambda y: f"={R('rev', y)}*{R('Assumptions.dc_pct', y)}", note=n)
    h, n = IS("us-gaap_GeneralAndAdministrativeExpense")
    sh.row("oh", "Operating overhead expense", hist=h,
           fc=lambda y: f"={R('rev', y)}*{R('Assumptions.oh_pct', y)}+{R('Assumptions.restr', y)}",
           note=n + "; forecast = % of revenue ex-restructuring + Pace charges")
    sh.row("sga", "Total selling and administrative expense", hist=total(["dc", "oh"]), fc=total(["dc", "oh"]), total=True)
    h, n = IS("us-gaap_InterestIncomeExpenseNonoperatingNet")
    sh.row("int", "Interest (income) expense, net", hist=h, fc=link("BS.net_int"), note=n + "; forecast from debt and cash schedule (BS sheet)")
    h, n = IS("us-gaap_OtherNonoperatingIncomeExpense")
    sh.row("oth", "Other (income) expense, net", hist=h, fc=link("Assumptions.oth"), note=n)
    pbt = lambda y: f"={R('gp', y)}-{R('sga', y)}-{R('int', y)}-{R('oth', y)}"
    sh.row("pbt", "Income before income taxes", hist=pbt, fc=pbt, total=True)
    h, n = IS("us-gaap_IncomeTaxExpenseBenefit")
    sh.row("tax", "Income tax expense", hist=h, fc=lambda y: f"={R('pbt', y)}*{R('Assumptions.tax', y)}", note=n)
    sh.row("etr", "Effective tax rate", "%", hist=ratio("tax", "pbt"), fc=ratio("tax", "pbt"), fmt=PCT, indent=1)
    sh.row("ni", "NET INCOME", hist=subtract("pbt", "tax"), fc=subtract("pbt", "tax"), total=True)
    sh.blank()

    sh.section("Per share")
    h, n = IS("us-gaap_WeightedAverageNumberOfSharesOutstandingBasic")
    sh.row("basic", "Weighted average basic shares", "m", hist=h, fmt=NUM1, note=n + "; forecast = prior x (1 + issuance %) - buyback $ / repurchase price (Assumptions)",
           fc=lambda y: f"={R('basic', y - 1)}*(1+{R('Assumptions.share_iss', y)})-{R('Assumptions.buyback', y)}/{R('Assumptions.bb_price', y)}")
    h, n = IS("us-gaap_WeightedAverageNumberOfDilutedSharesOutstanding")
    sh.row("diluted", "Weighted average diluted shares", "m", hist=h, fmt=NUM1, note=n + "; forecast = basic + FY2026 dilution",
           fc=lambda y: f"={R('basic', y)}+({R('diluted', 2026, True)}-{R('basic', 2026, True)})")
    sh.row("eps_b", "EPS - basic", "$", hist=ratio("ni", "basic"), fc=ratio("ni", "basic"), fmt=USD2)
    sh.row("eps_d", "EPS - diluted", "$", hist=ratio("ni", "diluted"), fc=ratio("ni", "diluted"), fmt=USD2, bold=True)
    h, rows = sup("Dividends declared per common share (USD)")
    sh.row("dps", "Dividends declared per share", "$", hist={y: v for y, v in h.items() if y in HIST}, fmt=USD2,
           fc=lambda y: f"={R('dps', y - 1)}*(1+{R('Assumptions.dps_g', y)})",
           note="XBRL us-gaap:CommonStockDividendsPerShareDeclared (equity statement); FY22 from FY24 10-K, FY23 from FY25 10-K")
    sh.row("div", "Dividends paid", hist=lambda y: f"=-{R('CF.div', y)}", fc=lambda y: f"={R('dps', y)}*{R('basic', y)}",
           note="Forecast = DPS x basic shares (ignores the one-quarter lag between declaration and payment).")
    sh.row("payout", "Dividend payout (dividends paid / net income)", "%", hist=ratio("div", "ni"), fc=ratio("div", "ni"), fmt=PCT, indent=1)
    sh.blank()

    sh.section("Memo: restructuring, EBIT and EBITDA")
    restr = {loc: sup("Severance and restructuring charges", loc)[0] for loc in ("CostOfSales", "GeneralAndAdministrative")}
    for key, loc, label in (("restr_cogs", "CostOfSales", "in cost of sales"), ("restr_oh", "GeneralAndAdministrative", "in operating overhead")):
        vals = {y: restr[loc].get(y, 0.0) for y in HIST}
        sh.row(key, f"Severance / restructuring charges {label}", hist=vals, indent=1,
               fc=link("Assumptions.restr") if key == "restr_oh" else None,
               note="XBRL us-gaap:RestructuringCharges [IncomeStatementLocationAxis] (FY24 and FY26 from FY26 10-K). "
                    "No charge tagged for FY22, FY23, FY25: shown as 0." + (" Forecast: all Pace charges in overhead." if key == "restr_oh" else ""))
    sh.row("restr_tot", "Total severance / restructuring charges", hist=total(["restr_cogs", "restr_oh"]), fc=total(["restr_cogs", "restr_oh"]))
    sh.row("da", "Depreciation and amortization", hist=lambda y: f"={R('CF.da', y)}+{R('CF.dep', y)}", fc=link("CF.da"),
           note="From cash flow statement. FY2022: depreciation only (amortization sits in 'Amortization, impairment and other').")
    ebit = lambda y: f"={R('pbt', y)}+{R('int', y)}"
    sh.row("ebit", "EBIT (Nike definition: pre-tax income + net interest expense)", hist=ebit, fc=ebit, bold=True)
    sh.row("ebit_m", "EBIT margin", "%", hist=ratio("ebit", "rev"), fc=ratio("ebit", "rev"), fmt=PCT, indent=1)
    sh.row("ebit_adj", "EBIT excluding restructuring", hist=total(["ebit", "restr_tot"]), fc=total(["ebit", "restr_tot"]))
    sh.row("ebit_adj_m", "EBIT margin excluding restructuring", "%", hist=ratio("ebit_adj", "rev"), fc=ratio("ebit_adj", "rev"), fmt=PCT, indent=1)
    sh.row("ebitda", "EBITDA", hist=total(["ebit", "da"]), fc=total(["ebit", "da"]), bold=True)
    sh.row("ebitda_m", "EBITDA margin", "%", hist=ratio("ebitda", "rev"), fc=ratio("ebitda", "rev"), fmt=PCT, indent=1)
    sh.row("net_m", "Net margin", "%", hist=ratio("ni", "rev"), fc=ratio("ni", "rev"), fmt=PCT, indent=1)
    return sh


def build_bs():
    sh = Sheet("BS", "NIKE, Inc. - Balance sheet (May 31) and supporting schedules", hist_years=BS_HIST,
               sources=column_sources("BS"), show_2021=True)
    B = lambda lid: stmt("BS", lid)

    def line(key, lid, label, fc, note_extra="", **kw):
        h, n = B(lid)
        sh.row(key, label, hist=h, fc=fc, note=n + (f"; {note_extra}" if note_extra else ""), **kw)

    flat = "forecast held flat"
    sh.section("Assets")
    line("cash", "us-gaap_CashAndCashEquivalentsAtCarryingValue", "Cash and equivalents", link("CF.end_cash"), "forecast = CF ending cash")
    line("sti", "us-gaap_DebtSecuritiesAvailableForSaleExcludingAccruedInterestCurrent", "Short-term investments", prior("sti"), flat)
    line("ar", "us-gaap_AccountsReceivableNetCurrent", "Accounts receivable, net",
         lambda y: f"={R('IS.rev', y)}*{R('Assumptions.dso', y)}/365", "forecast = revenue x DSO / 365")
    line("inv", "us-gaap_InventoryFinishedGoodsNetOfReserves", "Inventories",
         lambda y: f"={R('IS.cogs', y)}*{R('Assumptions.dio', y)}/365", "forecast = COGS x DIO / 365")
    line("prepaid", "us-gaap_PrepaidExpenseAndOtherAssetsCurrent", "Prepaid expenses and other current assets",
         lambda y: f"={R('IS.rev', y)}*{R('Assumptions.prepaid_pct', y)}", "forecast = % of revenue")
    ca = ["cash", "sti", "ar", "inv", "prepaid"]
    sh.row("tca", "Total current assets", hist=total(ca), fc=total(ca), total=True)
    line("ppe", "us-gaap_PropertyPlantAndEquipmentNet", "Property, plant and equipment, net",
         lambda y: f"={R('ppe', y - 1)}-{R('CF.capex', y)}-{R('CF.da', y)}", "forecast = prior + capex - D&A")
    line("rou", "us-gaap_OperatingLeaseRightOfUseAsset", "Operating lease right-of-use assets, net", prior("rou"), flat)
    line("intang", "us-gaap_IntangibleAssetsNetExcludingGoodwill", "Identifiable intangible assets, net", prior("intang"), flat)
    line("gw", "us-gaap_Goodwill", "Goodwill", prior("gw"), flat)
    line("dta", "us-gaap_DeferredIncomeTaxesAndOtherAssetsNoncurrent", "Deferred income taxes and other assets", prior("dta"), flat)
    ta = ["tca", "ppe", "rou", "intang", "gw", "dta"]
    sh.row("ta", "TOTAL ASSETS", hist=total(ta), fc=total(ta), total=True)
    sh.blank()

    sh.section("Liabilities and shareholders' equity")
    line("cur_ltd", "us-gaap_LongTermDebtCurrent", "Current portion of long-term debt", link("cur_port"), "forecast = next year's maturities")
    line("notes", "us-gaap_ShortTermBorrowings", "Notes payable", prior("notes"),
         "not on the face of the FY25/FY26 balance sheets (folded into accrued liabilities); forecast = prior (nil)")
    line("ap", "us-gaap_AccountsPayableCurrent", "Accounts payable",
         lambda y: f"={R('IS.cogs', y)}*{R('Assumptions.dpo', y)}/365", "forecast = COGS x DPO / 365")
    line("cur_lease", "us-gaap_OperatingLeaseLiabilityCurrent", "Current portion of operating lease liabilities", prior("cur_lease"), flat)
    line("accrued", "us-gaap_AccruedLiabilitiesCurrent", "Accrued liabilities",
         lambda y: f"={R('IS.rev', y)}*{R('Assumptions.accrued_pct', y)}", "forecast = % of revenue")
    line("tax_pay", "us-gaap_AccruedIncomeTaxesCurrent", "Income taxes payable", prior("tax_pay"), flat)
    cl = ["cur_ltd", "notes", "ap", "cur_lease", "accrued", "tax_pay"]
    sh.row("tcl", "Total current liabilities", hist=total(cl), fc=total(cl), total=True)
    line("ltd", "us-gaap_LongTermDebtNoncurrent", "Long-term debt", link("noncur"), "forecast from debt schedule below")
    line("lease", "us-gaap_OperatingLeaseLiabilityNoncurrent", "Operating lease liabilities", prior("lease"), flat)
    line("dtl", "us-gaap_DeferredIncomeTaxesAndOtherLiabilitiesNoncurrent", "Deferred income taxes and other liabilities", prior("dtl"), flat)
    tl = ["tcl", "ltd", "lease", "dtl"]
    sh.row("tl", "Total liabilities (not shown on Nike's balance sheet)", hist=total(tl), fc=total(tl), bold=True)
    line("pref", "us-gaap_TemporaryEquityCarryingAmountAttributableToParent", "Redeemable preferred stock", prior("pref"), flat)
    line("cs_a", "us-gaap_CommonStockValue[us-gaap:CommonClassAMember]", "Common stock at stated value - Class A", prior("cs_a"), flat, indent=1)
    line("cs_b", "us-gaap_CommonStockValue[us-gaap:CommonClassBMember]", "Common stock at stated value - Class B", prior("cs_b"), flat, indent=1)
    line("apic", "us-gaap_AdditionalPaidInCapitalCommonStock", "Capital in excess of stated value",
         lambda y: f"={R('apic', y - 1)}+{R('CF.sbc', y)}+{R('CF.options', y)}", "forecast = prior + SBC + option proceeds", indent=1)
    line("aoci", "us-gaap_AccumulatedOtherComprehensiveIncomeLossNetOfTax", "Accumulated other comprehensive income (loss)", prior("aoci"),
         flat + " (no FX or hedge forecast)", indent=1)
    line("re", "us-gaap_RetainedEarningsAccumulatedDeficit", "Retained earnings (deficit)",
         lambda y: f"={R('re', y - 1)}+{R('IS.ni', y)}+{R('CF.div', y)}+{R('CF.buyback', y)}",
         "forecast = prior + net income - dividends - buybacks (all repurchases charged here)", indent=1)
    eq = ["cs_a", "cs_b", "apic", "aoci", "re"]
    sh.row("eq", "Total shareholders' equity", hist=total(eq), fc=total(eq), total=True)
    sh.row("tle", "TOTAL LIABILITIES AND SHAREHOLDERS' EQUITY", hist=total(["tl", "pref", "eq"]), fc=total(["tl", "pref", "eq"]), total=True)
    chk = lambda y: f"=ROUND({R('ta', y)}-{R('tle', y)},3)"
    sh.row("bal", "Balance check: assets - liabilities and equity (must be 0)", hist=chk, fc=chk, bold=True)
    sh.blank()

    # Debt and interest schedule. FY2026 column holds opening balances; SUMPRODUCT over the bond table on Assumptions.
    only26 = lambda f: (lambda y: f(y) if y == 2026 else None)
    sh.section("Debt and interest schedule (senior notes; FY2026 = opening balances)")
    fy = lambda y: f"=VALUE(MID({COL[y]}$4,3,4))"
    sh.row("fy", "Fiscal year (from header)", "#", hist=only26(fy), fc=fy, fmt="0")
    face_exist = lambda y: f"=SUMPRODUCT({DEBT('C')}*({DEBT('F')}>{R('fy', y)}))"
    sh.row("face_exist", "Existing notes outstanding, end of year (face)", hist=only26(face_exist), fc=face_exist)
    sh.row("mat", "Maturities of existing notes", fc=lambda y: f"=SUMPRODUCT({DEBT('C')}*({DEBT('F')}={R('fy', y)}))",
           note="Matches the FY26 10-K maturity table: FY27 $2,000m, FY30 $1,500m.")
    sh.row("iss", "New notes issued (refinancing)", fc=lambda y: f"={R('mat', y)}*{S('refi_pct')}")
    sh.row("new_debt", "New notes outstanding, end of year", fc=lambda y: f"={R('new_debt', y - 1)}+{R('iss', y)}",
           note="New notes assumed long-dated (no maturities inside the forecast).")
    sh.row("face", "Total notes outstanding (face)", hist=only26(total(["face_exist", "new_debt"])), fc=total(["face_exist", "new_debt"]), bold=True)
    sh.row("carry_adj", "Unamortized discount and issuance costs", hist=only26(lambda y: f"={R('cur_ltd', y)}+{R('ltd', y)}-{R('face', y)}"),
           fc=prior("carry_adj"), note="FY2026 carrying value - face; held constant (amortization ignored, < $10m a year).")
    sh.row("cur_port", "Current portion (next year's maturities)",
           fc=lambda y: f"=SUMPRODUCT({DEBT('C')}*({DEBT('F')}={R('fy', y)}+1))")
    sh.row("noncur", "Long-term debt, non-current (carrying)", fc=lambda y: f"={R('face', y)}+{R('carry_adj', y)}-{R('cur_port', y)}")
    int_exist = lambda y: f"=SUMPRODUCT({DEBT('C')},{DEBT('D')},({DEBT('F')}>{R('fy', y)})+0.5*({DEBT('F')}={R('fy', y)}))"
    sh.row("int_exist", "Coupon interest on existing notes", hist=only26(int_exist), fc=int_exist,
           note="Maturing notes accrue half a year of interest (mid-year convention).")
    sh.row("int_new", "Interest on new notes", fc=lambda y: f"=({R('new_debt', y - 1)}+0.5*{R('iss', y)})*{S('new_debt_rate')}")
    sh.row("int_exp", "Total interest expense", fc=total(["int_exist", "int_new"]), bold=True)
    sh.row("int_inc", "Interest income on beginning cash + short-term investments",
           fc=lambda y: f"=({R('cash', y - 1)}+{R('sti', y - 1)})*{S('cash_yield')}",
           note="Beginning balances avoid a circular reference (interest -> net income -> cash -> interest).")
    sh.row("net_int", "Interest (income) expense, net -> IS", fc=subtract("int_exp", "int_inc"), bold=True)
    return sh


def build_cf():
    sh = Sheet("CF", "NIKE, Inc. - Cash flow statement", sources=column_sources("CF"))
    C = lambda lid: stmt("CF", lid)
    nf = "not forecast"

    def line(key, lid, label, fc, note_extra="", **kw):
        h, n = C(lid)
        sh.row(key, label, hist=h, fc=fc, note=n + (f"; {note_extra}" if note_extra else ""), **kw)

    sh.section("Cash provided (used) by operations")
    line("ni", "us-gaap_NetIncomeLoss", "Net income", link("IS.ni"))
    line("dep", "us-gaap_Depreciation", "Depreciation (FY2022 layout)", None, "FY2022 only; FY2023+ reported as one D&A line", indent=1)
    line("amort", "nke_AmortizationImpairmentAndOther", "Amortization, impairment and other (FY2022 layout)", None, "FY2022 only", indent=1)
    line("da", "us-gaap_DepreciationDepletionAndAmortization", "Depreciation and amortization",
         lambda y: f"={R('IS.rev', y)}*{R('Assumptions.da_pct', y)}", "forecast = % of revenue", indent=1)
    line("dtax", "us-gaap_DeferredIncomeTaxExpenseBenefit", "Deferred income taxes", None, nf, indent=1)
    line("sbc", "us-gaap_ShareBasedCompensation", "Stock-based compensation",
         lambda y: f"={R('IS.rev', y)}*{R('Assumptions.sbc_pct', y)}", "forecast = % of revenue", indent=1)
    line("impair", "us-gaap_AssetImpairmentCharges", "Impairment and other", None, nf, indent=1)
    line("fxadj", "us-gaap_ForeignCurrencyTransactionGainLossUnrealized", "Net foreign currency adjustments", None, nf, indent=1)
    line("d_ar", "us-gaap_IncreaseDecreaseInAccountsReceivable", "(Increase) decrease in accounts receivable",
         lambda y: f"=-({R('BS.ar', y)}-{R('BS.ar', y - 1)})", indent=1)
    line("d_inv", "us-gaap_IncreaseDecreaseInInventories", "(Increase) decrease in inventories",
         lambda y: f"=-({R('BS.inv', y)}-{R('BS.inv', y - 1)})", indent=1)
    other_a = ["prepaid", "rou", "intang", "gw", "dta"]
    line("d_oa", "us-gaap_IncreaseDecreaseInPrepaidDeferredExpenseAndOtherAssets",
         "(Increase) decrease in prepaid expenses, ROU and other assets",
         lambda y: "=-(" + "+".join(f"{R('BS.' + k, y)}-{R('BS.' + k, y - 1)}" for k in other_a) + ")", indent=1,
         note_extra="forecast = -change in prepaid, ROU, intangibles, goodwill, DTA and other")
    other_l = ["ap", "accrued", "tax_pay", "cur_lease", "lease", "dtl"]
    line("d_ol", "us-gaap_IncreaseDecreaseInAccountsPayableAndOtherOperatingLiabilities",
         "Increase (decrease) in accounts payable, accrued, lease and other liabilities",
         lambda y: "=" + "+".join(f"{R('BS.' + k, y)}-{R('BS.' + k, y - 1)}" for k in other_l), indent=1,
         note_extra="forecast = change in AP, accrued, taxes payable, lease liabilities, DTL and other")
    cfo_rows = lambda y: f"=SUM({R('ni', y)}:{R('d_ol', y)})"
    sh.row("cfo", "Cash provided (used) by operations", hist=cfo_rows, fc=cfo_rows, total=True)
    sh.blank()

    sh.section("Cash provided (used) by investing activities")
    line("sti_buy", "us-gaap_PaymentsToAcquireAvailableForSaleSecuritiesDebt", "Purchases of short-term investments",
         lambda y: f"=-({R('BS.sti', y)}-{R('BS.sti', y - 1)})", "forecast = net change in short-term investments", indent=1)
    line("sti_mat", "us-gaap_ProceedsFromMaturitiesPrepaymentsAndCallsOfAvailableForSaleSecurities", "Maturities of short-term investments",
         None, "forecast netted in the purchases line", indent=1)
    line("sti_sale", "us-gaap_ProceedsFromSaleOfAvailableForSaleSecuritiesDebt", "Sales of short-term investments", None,
         "forecast netted in the purchases line", indent=1)
    line("capex", "us-gaap_PaymentsToAcquirePropertyPlantAndEquipment", "Additions to property, plant and equipment",
         lambda y: f"=-{R('IS.rev', y)}*{R('Assumptions.capex_pct', y)}", "forecast = % of revenue", indent=1)
    line("oth_inv", "us-gaap_PaymentsForProceedsFromOtherInvestingActivities", "Other investing activities", None, nf, indent=1)
    cfi_rows = lambda y: f"=SUM({R('sti_buy', y)}:{R('oth_inv', y)})"
    sh.row("cfi", "Cash provided (used) by investing activities", hist=cfi_rows, fc=cfi_rows, total=True)
    sh.blank()

    sh.section("Cash provided (used) by financing activities")
    sh.row("debt_iss", "Proceeds from long-term debt issuance", fc=link("BS.iss"), indent=1,
           note="Model line: no issuance in FY2022-26 so no reported line; forecast from debt schedule")
    line("repay", "us-gaap_RepaymentsOfDebt", "Repayment of borrowings", lambda y: f"=-{R('BS.mat', y)}", "forecast = maturities", indent=1)
    line("d_notes", "us-gaap_ProceedsFromRepaymentsOfNotesPayable", "Increase (decrease) in notes payable, net",
         lambda y: f"={R('BS.notes', y)}-{R('BS.notes', y - 1)}", "line dropped after FY2023 layout", indent=1)
    line("options", "us-gaap_ProceedsFromIssuanceOrSaleOfEquity", "Proceeds from exercise of stock options and other stock issuances",
         link("Assumptions.options"), indent=1)
    line("buyback", "us-gaap_PaymentsForRepurchaseOfCommonStock", "Repurchase of common stock",
         lambda y: f"=-{R('Assumptions.buyback', y)}", indent=1)
    line("div", "us-gaap_PaymentsOfDividendsCommonStock", "Dividends - common and preferred", lambda y: f"=-{R('IS.div', y)}", indent=1)
    line("oth_fin", "us-gaap_ProceedsFromPaymentsForOtherFinancingActivities", "Other financing activities", None, nf, indent=1)
    cff_rows = lambda y: f"=SUM({R('debt_iss', y)}:{R('oth_fin', y)})"
    sh.row("cff", "Cash provided (used) by financing activities", hist=cff_rows, fc=cff_rows, total=True)
    sh.blank()

    line("fx", "us-gaap_EffectOfExchangeRateOnCashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents",
         "Effect of exchange rate changes on cash and equivalents", None, nf)
    chg = lambda y: f"={R('cfo', y)}+{R('cfi', y)}+{R('cff', y)}+{R('fx', y)}"
    sh.row("chg", "Net increase (decrease) in cash and equivalents", hist=chg, fc=chg, bold=True)
    line("beg_cash", "us-gaap_CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents[start]",
         "Cash and equivalents, beginning of year", prior("end_cash"))
    end = lambda y: f"={R('beg_cash', y)}+{R('chg', y)}"
    sh.row("end_cash", "CASH AND EQUIVALENTS, END OF YEAR", hist=end, fc=end, total=True,
           note="Equals BS cash in every historical year (no restricted cash difference); feeds BS cash in the forecast.")
    sh.blank()

    sh.section("Supplemental and memo")
    line("int_paid", "us-gaap_InterestPaidNet", "Cash paid for interest, net of capitalized interest", None, "history only", indent=1)
    line("tax_paid", "us-gaap_IncomeTaxesPaid", "Cash paid for income taxes", None,
         "on the face of the cash flow statement through FY2023 only (moved to the tax note)", indent=1)
    line("ncapex", "us-gaap_CapitalExpendituresIncurredButNotYetPaid", "Non-cash additions to PP&E", None, "history only", indent=1)
    sh.row("fcf", "Free cash flow (CFO - capex)", hist=total(["cfo", "capex"]), fc=total(["cfo", "capex"]), bold=True)
    return sh


def build_ratios():
    sh = Sheet("Ratios", "NIKE, Inc. - Ratio analysis")
    avg = lambda key, y: f"({R(key, y)}+{R(key, y - 1)})/2"
    sh.section("Growth")
    for k, label in (("IS.rev_g", "Revenue growth"), ("Segments.nb_g", "NIKE Brand revenue growth"),
                     ("Segments.nb_d_g", "NIKE Brand Direct growth"), ("Segments.nb_w_g", "NIKE Brand Wholesale growth")):
        sh.row(k.split(".")[1], label, "%", hist=link(k), fc=link(k), fmt=PCT)
    sh.row("eps_g", "Diluted EPS growth", "%", hist=growth_rate("IS.eps_d"), fc=growth_rate("IS.eps_d"), fmt=PCT)
    sh.blank()

    sh.section("Margins")
    for k, label in (("IS.gm", "Gross margin"), ("Assumptions.dc_pct", "Demand creation, % of revenue")):
        sh.row(k.split(".")[1], label, "%", hist=link(k), fc=link(k), fmt=PCT)
    sh.row("oh_pct", "Operating overhead (reported), % of revenue", "%", hist=ratio("IS.oh", "IS.rev"), fc=ratio("IS.oh", "IS.rev"), fmt=PCT)
    for k, label in (("IS.ebit_m", "EBIT margin"), ("IS.ebit_adj_m", "EBIT margin ex-restructuring"), ("IS.ebitda_m", "EBITDA margin"),
                     ("IS.net_m", "Net margin"), ("IS.etr", "Effective tax rate")):
        sh.row(k.split(".")[1], label, "%", hist=link(k), fc=link(k), fmt=PCT)
    sh.blank()

    sh.section("Returns (on average balances)")
    sh.row("roe", "Return on equity (net income / average equity)", "%",
           hist=lambda y: f"={R('IS.ni', y)}/({avg('BS.eq', y)})", fc=lambda y: f"={R('IS.ni', y)}/({avg('BS.eq', y)})", fmt=PCT)
    ic_keys = ["cur_ltd", "notes", "ltd", "cur_lease", "lease", "eq"]
    ic = lambda y: "=" + "+".join(R("BS." + k, y) for k in ic_keys) + f"-{R('BS.cash', y)}-{R('BS.sti', y)}"
    sh.row("ic", "Invested capital (debt + lease liabilities + equity - cash - STI)", hist=ic, fc=ic,
           hist_years=BS_HIST, note="Analyst definition; Nike's own reported ROIC uses a different method, so the two will differ.")
    sh.row("nopat", "NOPAT (EBIT x (1 - effective tax rate))", hist=lambda y: f"={R('IS.ebit', y)}*(1-{R('IS.etr', y)})",
           fc=lambda y: f"={R('IS.ebit', y)}*(1-{R('IS.etr', y)})")
    sh.row("roic", "ROIC (NOPAT / average invested capital)", "%",
           hist=lambda y: f"={R('nopat', y)}/({avg('ic', y)})", fc=lambda y: f"={R('nopat', y)}/({avg('ic', y)})", fmt=PCT)
    sh.blank()

    sh.section("Working capital (ending balances, 365 days)")
    dso = lambda y: f"={R('BS.ar', y)}/{R('IS.rev', y)}*365"
    dio = lambda y: f"={R('BS.inv', y)}/{R('IS.cogs', y)}*365"
    dpo = lambda y: f"={R('BS.ap', y)}/{R('IS.cogs', y)}*365"
    sh.row("dso", "DSO", "days", hist=dso, fc=dso, fmt=DAYS)
    sh.row("dio", "DIO", "days", hist=dio, fc=dio, fmt=DAYS)
    sh.row("dpo", "DPO", "days", hist=dpo, fc=dpo, fmt=DAYS)
    ccc = lambda y: f"={R('dso', y)}+{R('dio', y)}-{R('dpo', y)}"
    sh.row("ccc", "Cash conversion cycle", "days", hist=ccc, fc=ccc, fmt=DAYS, bold=True)
    sh.blank()

    sh.section("Leverage and coverage")
    debt = lambda y: f"={R('BS.cur_ltd', y)}+{R('BS.notes', y)}+{R('BS.ltd', y)}"
    sh.row("debt", "Total debt (notes payable + current and long-term debt)", hist=debt, fc=debt, hist_years=BS_HIST)
    nd = lambda y: f"={R('debt', y)}-{R('BS.cash', y)}-{R('BS.sti', y)}"
    sh.row("net_debt", "Net debt (negative = net cash)", hist=nd, fc=nd, hist_years=BS_HIST)
    sh.row("debt_ebitda", "Total debt / EBITDA", "x", hist=ratio("debt", "IS.ebitda"), fc=ratio("debt", "IS.ebitda"), fmt=MULT)
    sh.row("nd_ebitda", "Net debt / EBITDA", "x", hist=ratio("net_debt", "IS.ebitda"), fc=ratio("net_debt", "IS.ebitda"), fmt=MULT)
    sh.row("de", "Debt / equity", "x", hist=ratio("debt", "BS.eq"), fc=ratio("debt", "BS.eq"), fmt=MULT)
    sh.row("cover", "EBIT / interest expense", "x", hist=ratio("IS.ebit", "CF.int_paid"), fc=ratio("IS.ebit", "BS.int_exp"), fmt=MULT,
           note="History uses cash interest paid as a proxy (gross interest expense is not on the face of the income statement).")
    sh.blank()

    sh.section("Cash flow and capital returns")
    sh.row("fcf", "Free cash flow", hist=link("CF.fcf"), fc=link("CF.fcf"))
    sh.row("fcf_conv", "FCF / net income", "%", hist=ratio("fcf", "IS.ni"), fc=ratio("fcf", "IS.ni"), fmt=PCT)
    sh.row("capex_da", "Capex / D&A", "x", hist=lambda y: f"=-{R('CF.capex', y)}/{R('IS.da', y)}",
           fc=lambda y: f"=-{R('CF.capex', y)}/{R('IS.da', y)}", fmt=MULT)
    sh.row("payout", "Dividend payout", "%", hist=link("IS.payout"), fc=link("IS.payout"), fmt=PCT)
    tsr = lambda y: f"=-({R('CF.div', y)}+{R('CF.buyback', y)})/{R('IS.ni', y)}"
    sh.row("return_ni", "(Dividends + buybacks) / net income", "%", hist=tsr, fc=tsr, fmt=PCT)
    return sh


def build_dcf():
    sh = Sheet("DCF", "NIKE, Inc. - DCF valuation and WACC", c_label="Value")
    it = sh.items
    V = lambda k: S(k, "DCF")

    sh.section("Unlevered free cash flow (history for reference; FY2027E-FY2031E are valued)")
    nopat = lambda y: f"={R('ebit', y)}*(1-{R('tax', y)})"
    nwc = lambda y: (f"={R('BS.ar', y)}+{R('BS.inv', y)}+{R('BS.prepaid', y)}"
                     f"-{R('BS.ap', y)}-{R('BS.accrued', y)}")
    d_nwc = lambda y: f"=-({R('nwc', y)}-{R('nwc', y - 1)})"
    ufcf = total(["nopat", "da", "capex", "d_nwc"])
    from_2023 = lambda f: (lambda y: None if y < 2023 else f(y))
    sh.row("rev", "Revenues", hist=link("IS.rev"), fc=link("IS.rev"))
    sh.row("ebit", "EBIT", hist=link("IS.ebit"), fc=link("IS.ebit"),
           note="Nike definition (pre-tax income + net interest expense); after restructuring and other (income) expense.")
    sh.row("ebit_m", "EBIT margin", "%", hist=ratio("ebit", "rev"), fc=ratio("ebit", "rev"), fmt=PCT, indent=1)
    sh.row("tax", "Tax rate on EBIT", "%", hist=link("IS.etr"), fc=link("Assumptions.tax"), fmt=PCT, indent=1,
           note="Effective tax rate: reported (history), forecast assumption (FY2027E+).")
    sh.row("nopat", "NOPAT = EBIT x (1 - tax rate)", hist=nopat, fc=nopat)
    sh.row("da", "Plus: depreciation and amortization", hist=link("IS.da"), fc=link("IS.da"), indent=1)
    sh.row("capex", "Less: capital expenditures", hist=link("CF.capex"), fc=link("CF.capex"), indent=1)
    sh.row("nwc", "Operating working capital (AR + inventories + prepaid - AP - accrued)", hist=nwc, fc=nwc, indent=1,
           note="Excludes cash, debt, taxes payable and lease liabilities (operating lease cost is already in EBIT under US GAAP).")
    sh.row("d_nwc", "Less: increase in operating working capital", hist=from_2023(d_nwc), fc=d_nwc, indent=1)
    sh.row("ufcf", "Unlevered free cash flow", hist=from_2023(ufcf), fc=ufcf, total=True,
           note="Stock-based compensation is not added back: it is treated as a real cost (EBIT already deducts it).")
    sh.row("sbc", "Memo: stock-based compensation (not added back)", hist=link("CF.sbc"), fc=link("CF.sbc"), indent=1)
    sh.blank()

    sh.section("Discounting (mid-period convention, to the valuation date on Assumptions)")
    sh.row("fy_end", "Fiscal year end", "date", fc=lambda y: f"=DATE({R('BS.fy', y)},5,31)", fmt="yyyy-mm-dd")
    sh.row("frac", "Share of the fiscal year after the valuation date", "%",
           fc=lambda y: f"=MAX(0,MIN(1,({R('fy_end', y)}-{S('price_date')})/365))", fmt=PCT,
           note="FY2027E: only the part of the year after 2026-10-30 is valued (cash flows assumed even through the year).")
    sh.row("pre_val", "Less: cash already received before the valuation date (IEEPA refund, in 2026-08-31 cash)",
           fc=lambda y: f"=-{S('ieepa')}*({R('BS.fy', y)}={R('BS.fy', 2027, True)})", indent=1,
           note="The $684m IEEPA receivable was collected in Q1 FY27, so it is already in the net cash used below; removed so it is not counted twice.")
    sh.row("ufcf_v", "Unlevered FCF valued = (FCF - cash already received) x share of year",
           fc=lambda y: f"=({R('ufcf', y)}+{R('pre_val', y)})*{R('frac', y)}")
    sh.row("period", "Discount period, years (mid-point of each cash flow)", "years",
           fc=lambda y: f"=SUM({R('frac', 2027, True)}:{R('frac', y)})-{R('frac', y)}/2", fmt="0.000",
           note="FY2027E stub: half of the 0.58-year stub (~0.29 years), not 0.5. Later years: stub + full years before + 0.5.")
    sh.row("df", "Discount factor at WACC", "x", fc=lambda y: f"=1/(1+{V('wacc')})^{R('period', y)}", fmt="0.0000")
    sh.row("pv", "Present value of unlevered FCF", fc=lambda y: f"={R('ufcf_v', y)}*{R('df', y)}", bold=True)
    sh.blank()

    sh.section("Balance sheet and shares at 2026-08-31 (Q1 FY27 10-Q; latest before the valuation date)")
    for key, label, st, concept in (
            ("q1_cash", "Cash and equivalents", "BS", "us-gaap_CashAndCashEquivalentsAtCarryingValue"),
            ("q1_sti", "Short-term investments", "BS", "us-gaap_DebtSecuritiesAvailableForSaleExcludingAccruedInterestCurrent"),
            ("q1_curltd", "Current portion of long-term debt", "BS", "us-gaap_LongTermDebtCurrent"),
            ("q1_ltd", "Long-term debt", "BS", "us-gaap_LongTermDebtNoncurrent"),
            ("q1_lease_cur", "Memo: operating lease liabilities, current (not treated as debt)", "BS", "us-gaap_OperatingLeaseLiabilityCurrent"),
            ("q1_lease_nc", "Memo: operating lease liabilities, non-current (not treated as debt)", "BS", "us-gaap_OperatingLeaseLiabilityNoncurrent"),
            ("q1_basic", "Q1 FY27 weighted average basic shares", "IS", "us-gaap_WeightedAverageNumberOfSharesOutstandingBasic"),
            ("q1_diluted", "Q1 FY27 weighted average diluted shares", "IS", "us-gaap_WeightedAverageNumberOfDilutedSharesOutstanding")):
        v, src = q1(st, concept)
        it.append(Scalar(key, label, "m" if st == "IS" else "$m", v, fmt=NUM1 if st == "IS" else NUM, note=src))
    it.append(Scalar("debt", "Total debt (current + long-term, carrying value)", "$m",
                     lambda: f"={V('q1_curltd')}+{V('q1_ltd')}", note="No notes payable or commercial paper at 2026-08-31."))
    it.append(Scalar("net_debt", "Net debt = debt - cash - short-term investments", "$m",
                     lambda: f"={V('debt')}-{V('q1_cash')}-{V('q1_sti')}", note="Negative = net cash."))
    it.append(Scalar("dil_shares", "Diluted shares for per-share value", "m",
                     lambda: f"={S('shares_out')}+{V('q1_diluted')}-{V('q1_basic')}", fmt=NUM1,
                     note="Shares outstanding at 2026-09-28 (Assumptions) + Q1 dilutive securities (diluted - basic weighted shares)."))
    sh.blank()

    sh.section("Cost of equity (CAPM)")
    it.append(Scalar("rf", "Risk-free rate (10-year Treasury)", "%", lambda: f"={S('rf')}", fmt=PCT2))
    it.append(Scalar("beta_raw", "Raw beta: SLOPE of 60 monthly NKE returns on S&P 500 TR returns", "x",
                     lambda: f"=SLOPE(E{REG[('DCF', 'beta_first')]}:E{REG[('DCF', 'beta_last')]},"
                             f"F{REG[('DCF', 'beta_first')]}:F{REG[('DCF', 'beta_last')]})", fmt="0.000",
                     note="Regression data at the bottom of this sheet (Nov 2021 - Oct 2026 month-end returns, dividends reinvested)."))
    beta_chk, beta_src = mkt("beta_raw_check")
    it.append(Scalar("beta_check", "Raw beta cross-check (Python OLS, market_data.py)", "x", beta_chk, fmt="0.000",
                     note="Independent calculation on the same returns; Checks sheet flags any difference above 0.001."))
    it.append(Scalar("beta_adj", "Adjusted beta = w x raw + (1 - w) x 1.0", "x",
                     lambda: f"={S('blume')}*{V('beta_raw')}+(1-{S('blume')})", fmt="0.000"))
    it.append(Scalar("erp", "Equity risk premium", "%", lambda: f"={S('erp')}", fmt=PCT2))
    it.append(Scalar("ke", "Cost of equity = Rf + adjusted beta x ERP", "%",
                     lambda: f"={V('rf')}+{V('beta_adj')}*{V('erp')}", fmt=PCT2))
    sh.blank()

    sh.section("Cost of debt")
    it.append(Scalar("kd_pre", "Pre-tax cost of debt (Treasury + A2/A spread)", "%", lambda: f"={S('kd_pre')}", fmt=PCT2))
    it.append(Scalar("tax_w", "Tax rate (FY2027E effective rate)", "%", lambda: f"={R('Assumptions.tax', 2027)}", fmt=PCT2))
    it.append(Scalar("kd_post", "After-tax cost of debt", "%", lambda: f"={V('kd_pre')}*(1-{V('tax_w')})", fmt=PCT2))
    sh.blank()

    sh.section("WACC (market value of equity, book value of debt)")
    it.append(Scalar("mcap", "Market value of equity = price x shares outstanding", "$m",
                     lambda: f"={S('price')}*{S('shares_out')}"))
    it.append(Scalar("we", "Weight of equity", "%", lambda: f"={V('mcap')}/({V('mcap')}+{V('debt')})", fmt=PCT))
    it.append(Scalar("wd", "Weight of debt", "%", lambda: f"={V('debt')}/({V('mcap')}+{V('debt')})", fmt=PCT,
                     note="Book debt approximates market value for investment-grade notes; leases excluded (consistent with US GAAP EBIT)."))
    it.append(Scalar("wacc", "WACC", "%", lambda: f"={V('we')}*{V('ke')}+{V('wd')}*{V('kd_post')}", fmt=PCT2))
    sh.blank()

    sh.section("Enterprise value, equity value and value per share")
    it.append(Scalar("sum_pv", "Sum of PV of unlevered FCF, FY2027E (stub) - FY2031E", "$m",
                     lambda: f"=SUM({R('pv', 2027)}:{R('pv', 2031)})"))
    it.append(Scalar("tv_period", "Terminal value discount period (end of FY2031), years", "years",
                     lambda: f"={R('period', 2031)}+{R('frac', 2031)}/2", fmt="0.000"))
    it.append(Scalar("ebitda31", "FY2031E EBITDA", "$m", lambda: f"={R('IS.ebitda', 2031)}"))
    for m, name in (("g", "Perpetuity growth method"), ("m", "Exit multiple method")):
        it.append(Text(name, italic=True))
        if m == "g":
            it.append(Scalar("tv_g", "Terminal value = FCF FY2031E x (1 + g) / (WACC - g)", "$m",
                             lambda: f"={R('ufcf', 2031)}*(1+{S('tg')})/({V('wacc')}-{S('tg')})"))
        else:
            it.append(Scalar("tv_m", "Terminal value = FY2031E EBITDA x exit multiple", "$m",
                             lambda: f"={V('ebitda31')}*{S('exit_mult')}"))
        it.append(Scalar(f"pv_tv_{m}", "PV of terminal value", "$m", lambda m=m: f"={V('tv_' + m)}/(1+{V('wacc')})^{V('tv_period')}"))
        it.append(Scalar(f"ev_{m}", "Enterprise value", "$m", lambda m=m: f"={V('sum_pv')}+{V('pv_tv_' + m)}"))
        it.append(Scalar(f"tvsh_{m}", "Terminal value share of EV", "%", lambda m=m: f"={V('pv_tv_' + m)}/{V('ev_' + m)}", fmt=PCT))
        it.append(Scalar(f"eq_{m}", "Equity value = EV - net debt", "$m", lambda m=m: f"={V('ev_' + m)}-{V('net_debt')}"))
        it.append(Scalar(f"px_{m}", "Value per share", "$/share", lambda m=m: f"={V('eq_' + m)}/{V('dil_shares')}", fmt=USD2,
                         key_input=False))
        it.append(Scalar(f"up_{m}", "Upside (downside) vs. current price", "%", lambda m=m: f"={V('px_' + m)}/{S('price')}-1", fmt=PCT))
    it.append(Text("Terminal values are discounted from the end of FY2031 under both methods (the mid-period convention is "
                   "applied to the FY2027E-FY2031E cash flows only). Net debt is at 2026-08-31; cash flows between then and the "
                   "valuation date are not separately adjusted (apart from the IEEPA refund above).", italic=True))
    sh.blank()

    sh.section("Terminal value cross-check: what each method implies about the other")
    it.append(Scalar("impl_mult", "Implied exit multiple from the perpetuity-growth terminal value (TV / FY2031E EBITDA)", "x",
                     lambda: f"={V('tv_g')}/{V('ebitda31')}", fmt=MULT, note="Compare with the exit multiple assumption on Assumptions."))
    it.append(Scalar("impl_g", "Implied perpetual growth from the exit-multiple terminal value = (TV x WACC - FCF) / (TV + FCF)", "%",
                     lambda: f"=({V('tv_m')}*{V('wacc')}-{R('ufcf', 2031)})/({V('tv_m')}+{R('ufcf', 2031)})", fmt=PCT2,
                     note="Solves TV = FCF FY2031E x (1 + g) / (WACC - g) for g. Compare with the terminal growth assumption."))
    it.append(Scalar("tv_gap", "Value per share: perpetuity growth less exit multiple", "$/share",
                     lambda: f"={V('px_g')}-{V('px_m')}", fmt=USD2))
    sh.blank()

    def grid_render(kind):
        def render(ws, r0):
            CUR["sheet"] = "DCF"
            row = lambda k: REG[("DCF", k)]
            if kind == "pg":
                title, head, step, col_fmt = "Value per share: WACC (rows) x terminal growth (columns)", S("tg"), S("g_step"), PCT2
            elif kind == "em":
                title, head, step, col_fmt = ("Value per share: WACC (rows) x exit EV/EBITDA multiple (columns)", S("exit_mult"),
                                              S("mult_step"), MULT)
            else:
                title, head, step, col_fmt = ("Value per share: WACC (rows) x FY2030E EBIT margin (columns), perpetuity growth method",
                                              f"$L${row('ebit_m')}", S("margin_step"), PCT)
            ws.cell(r0, 1, title).font = font(bold=True)
            h = r0 + 1
            ws.cell(h, 1, "Centre cell = model values; steps on Assumptions").font = font(GREY, italic=True, size=9)
            for j, k in enumerate(range(-2, 3)):
                c = ws.cell(h, 4 + j, f"={head}+({k})*{step}")
                c.font, c.number_format, c.fill = font(GREEN, bold=True), col_fmt, SECTION_FILL
            pv = f"SUMPRODUCT($I${row('ufcf_v')}:$M${row('ufcf_v')}/(1+$C{{r}})^$I${row('period')}:$M${row('period')})"
            for i, k in enumerate(range(-2, 3)):
                r = h + 1 + i
                w = f"$C{r}"
                c = ws.cell(r, 3, f"={V('wacc')}+({k})*{S('wacc_step')}")
                c.font, c.number_format, c.fill = font(GREEN, bold=True), PCT2, SECTION_FILL
                for j in range(5):
                    col = "DEFGH"[j]
                    if kind == "pg":
                        tv = f"$M${row('ufcf')}*(1+{col}${h})/({w}-{col}${h})/(1+{w})^{V('tv_period')}"
                    elif kind == "em":
                        tv = f"{V('ebitda31')}*{col}${h}/(1+{w})^{V('tv_period')}"
                    else:
                        # after-tax change in FY2030E and FY2031E EBIT, plus the same change carried into the terminal value
                        after_tax = lambda c: f"${c}${row('rev')}*(1-${c}${row('tax')})"
                        delta = (f"({col}${h}-$L${row('ebit_m')})*({after_tax('L')}*$L${row('frac')}/(1+{w})^$L${row('period')}"
                                 f"+{after_tax('M')}*$M${row('frac')}/(1+{w})^$M${row('period')}"
                                 f"+{after_tax('M')}*(1+{S('tg')})/({w}-{S('tg')})/(1+{w})^{V('tv_period')})")
                        tv = f"$M${row('ufcf')}*(1+{S('tg')})/({w}-{S('tg')})/(1+{w})^{V('tv_period')}+{delta}"
                    f = f"=({pv.format(r=r)}+{tv}-{V('net_debt')})/{V('dil_shares')}"
                    cell = ws.cell(r, 4 + j, f)
                    cell.font, cell.number_format = font(GREEN if "!" in f else BLACK, bold=(i == 2 and j == 2)), USD2
                    if i == 2 and j == 2:
                        cell.fill = YELLOW
            note = ("Each cell recomputes the full DCF (no data table): PV of FCF at that WACC + discounted terminal value - net debt, "
                    "/ diluted shares.")
            if kind == "mg":
                note += (" Margin change vs. the model's FY2030E EBIT margin is applied (after tax) to FY2030E and FY2031E EBIT and "
                         "carried into the terminal value; revenue, capex and working capital are unchanged.")
            ws.cell(h, 14, note).font = font(GREY, size=8)
        return render

    sh.section("Sensitivity: perpetuity growth method")
    it.append(Block(7, grid_render("pg"), {"pg_hdr": 1}))
    sh.section("Sensitivity: exit multiple method")
    it.append(Block(7, grid_render("em"), {"em_hdr": 1}))
    sh.section("Sensitivity: FY2030E EBIT margin (operating recovery) x WACC")
    it.append(Block(7, grid_render("mg"), {"mg_hdr": 1}))
    sh.blank()

    def beta_render(ws, r0):
        for col, text in ((1, "Month-end"), (3, "NKE adj. close"), (4, "S&P 500 TR"), (5, "NKE return"), (6, "S&P 500 TR return")):
            c = ws.cell(r0, col, text)
            c.font, c.fill = font(bold=True), SECTION_FILL
        ws.cell(r0, 14, "Adjusted closes (dividends reinvested). " + beta_src).font = font(GREY, size=8)
        for i, row in enumerate(BETA.itertuples()):
            r = r0 + 1 + i
            ws.cell(r, 1, row.month).font = font()
            for col, v in ((3, row.nke_adjclose), (4, row.sp500tr_adjclose)):
                c = ws.cell(r, col, round(float(v), 6))
                c.font, c.number_format = font(BLUE), "#,##0.00"
            if i > 0:
                for col, src in ((5, "C"), (6, "D")):
                    c = ws.cell(r, col, f"={src}{r}/{src}{r - 1}-1")
                    c.font, c.number_format = font(), PCT
    sh.section("Beta regression data: 61 month-ends, 60 monthly returns (Yahoo Finance)")
    it.append(Block(len(BETA) + 1, beta_render, {"beta_first": 2, "beta_last": len(BETA)}))
    return sh


COMPS = [  # ticker, name, accounting, reporting currency, fiscal year end
    ("NKE", "NIKE, Inc.", "US GAAP", "USD", "May 31"),
    ("ADS.DE", "adidas AG", "IFRS", "EUR", "Dec 31"),
    ("PUM.DE", "PUMA SE", "IFRS", "EUR", "Dec 31"),
    ("LULU", "lululemon athletica", "US GAAP", "USD", "~Feb 1 (52/53 wk)"),
    ("DECK", "Deckers Outdoor", "US GAAP", "USD", "Mar 31"),
    ("ONON", "On Holding AG", "IFRS", "CHF", "Dec 31"),
    ("UAA", "Under Armour", "US GAAP", "USD", "Mar 31"),
]
COMPS_NOTES = [
    "Lease treatment: IFRS 16 puts all leases on the balance sheet and moves lease cost below EBITDA, so for the IFRS peers (adidas, PUMA, On) "
    "EV includes lease liabilities. Under US GAAP, operating lease cost stays in EBITDA, so for Nike and the US GAAP peers EV excludes "
    "operating lease liabilities. Each EV/EBITDA is therefore internally consistent; the residual difference (IFRS EBITDA is higher by the "
    "full lease cost, EV by the lease liability) is not adjusted.",
    "Currency: multiples are computed in each company's reporting currency, so FX only matters where the share price and the accounts differ "
    "(On: USD price, CHF accounts, converted at ECB EUR/USD and EUR/CHF reference rates for 2026-10-30). USD market cap and EV rows are memo only.",
    "LTM periods differ by fiscal calendar (LTM ends Aug 2026 for Nike and Lululemon, Jun 2026 for On and Under Armour, Sep 2026 for "
    "adidas, PUMA and Deckers). No consensus forward multiples are used: no free, citable source is available.",
    "adidas: LTM EBITDA = LTM operating profit + FY2025 D&A from the 2025 Annual Report (EUR 1,135m, including impairment "
    "losses/reversals; adidas discloses no nine-month D&A, so the FY figure stands in for LTM). Cross-check: adjusted net borrowings / the "
    "reported leverage ratio (1.3x, rounded) gives a similar figure. Net income includes discontinued operations. "
    "Pensions (EUR 68m) are excluded from EV for consistency with the other peers.",
    "PUMA: D&A is not disclosed in the releases, so EBITDA is not available and EV/EBITDA is shown as N/M; P/E is N/M (net loss). Share count is the Q3 2026 weighted average "
    "(period-end shares are not stated). Non-controlling interests are not shown separately on the balance sheet (not in EV). "
    "Figures exclude PUMA United (discontinued operations) in all periods.",
    "Under Armour: Class A (UAA) and Class B (unlisted, convertible 1:1 into A) valued at the UAA price; Class C at the UA price. "
    "On: Class B voting shares have 1/10 of the nominal value and economic rights of a Class A share, so they count as 0.1 Class A shares. "
    "Lululemon: exchangeable shares are included (exchangeable 1:1 into common).",
    "Under Armour: LTM EBITDA and net income are negative, so EV/EBITDA and P/E are N/M. "
    "N/M = not meaningful (negative earnings or EBITDA, or not computable from disclosed figures); n/a = not disclosed. "
    "Peer statistics exclude Nike and ignore N/M. "
    "Hover over any blue input for its filing, page or XBRL tag.",
]


def build_comps():
    sh = Sheet("Comps", "Trading comparables - LTM multiples at the 2026-10-30 close", years=False,
               subtitle="Each company in its own reporting currency, millions (shares in millions, prices per share). "
                        "LTM = last fiscal year + current year-to-date - prior year-to-date.",
               widths={"A": 52, "B": 10, **{c: 14 for c in "CDEFGHI"}, "J": 110})
    tickers = [c[0] for c in COMPS]
    cols = {t: "CDEFGHI"[i] for i, t in enumerate(tickers)}
    peers = tickers[1:]
    lines = []  # (key, label, unit, fmt, {ticker: cell}, note, bold); cell = (value, source) input | "=formula" | text

    def peer(t, item, period="BS"):
        r = PEERS[(PEERS.ticker == t) & (PEERS.item == item) & (PEERS.period == period)]
        if r.empty:
            return None
        r = r.iloc[0]
        when = f"{r.start} to {r.end}" if isinstance(r.start, str) else f"at {r.end}"
        return float(r.value), f"{when} | {r.source}"

    def line(key, label, unit, fmt, cells, note="", bold=False):
        lines.append((key, label, unit, fmt, {t: v for t, v in cells.items() if v is not None}, note, bold))

    def every(formula, **override):
        return {t: override.get(t.replace(".", "_"), formula) for t in tickers}

    def section(text):
        lines.append((None, text, "", "", {}, "", True))

    # Nike's own Q1 FY27 / Q1 FY26 figures for its LTM
    nke = {}
    for item, st, concept in (("revenue", "IS", "us-gaap_RevenueFromContractWithCustomerExcludingAssessedTax"),
                              ("da", "CF", "us-gaap_DepreciationDepletionAndAmortization"),
                              ("net_income", "IS", "us-gaap_NetIncomeLoss")):
        for period, q in (("ytd", "Q1 FY2027"), ("ytd_prior", "Q1 FY2026")):
            nke[(item, period)] = q1(st, concept, q)
    for period, q in (("ytd", "Q1 FY2027"), ("ytd_prior", "Q1 FY2026")):
        pbt, src = q1("IS", "us-gaap_IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest", q)
        net_int, _ = q1("IS", "us-gaap_InterestIncomeExpenseNonoperatingNet", q)
        nke[("ebit", period)] = (pbt + net_int, f"Nike EBIT definition: pre-tax income {pbt:,.0f} + net interest (income) expense "
                                                f"{net_int:,.0f} (XBRL us-gaap:InterestIncomeExpenseNonoperatingNet). {src}")

    line("name", "Company", "", "@", {t: c[1] for t, c in zip(tickers, COMPS)}, bold=True)
    line("gaap", "Accounting standard", "", "@", {t: c[2] for t, c in zip(tickers, COMPS)})
    line("ccy", "Reporting currency", "", "@", {t: c[3] for t, c in zip(tickers, COMPS)})
    line("fye", "Fiscal year end", "", "@", {t: c[4] for t, c in zip(tickers, COMPS)})
    ltm_end = {"NKE": "2026-08-31", **{t: str(PEERS[(PEERS.ticker == t) & (PEERS.period == "YTD")].end.iloc[0]) for t in peers}}
    line("ltm_end", "LTM period ends", "", "@", ltm_end)

    section("Market value (closing prices on 2026-10-30)")
    line("price", "Share price (listing currency)", "per share", "#,##0.00",
         {"NKE": "={x_price}", **{t: mkt(f"price_{t}") for t in peers}},
         note="EUR for adidas and PUMA (Xetra); USD for the others (On's NYSE shares are priced in USD).")
    line("price2", "Share price: UA (Under Armour Class C)", "USD", "#,##0.00", {"UAA": mkt("price_UA")})
    line("sh1", "Shares: common / Class A", "m", NUM1,
         {"NKE": "={x_sha}", "ADS.DE": peer("ADS.DE", "shares"), "PUM.DE": peer("PUM.DE", "shares"), "LULU": peer("LULU", "shares"),
          "DECK": peer("DECK", "shares"), "ONON": peer("ONON", "shares_class_a"), "UAA": peer("UAA", "shares_class_a")})
    line("sh2", "Shares: Class B / exchangeable", "m", NUM1,
         {"NKE": "={x_shb}", "LULU": peer("LULU", "shares_exchangeable"), "ONON": peer("ONON", "shares_class_b"),
          "UAA": peer("UAA", "shares_class_b")})
    line("sh3", "Shares: Class C (UA)", "m", NUM1, {"UAA": peer("UAA", "shares_class_c")})
    line("econ", "Economic shares (Class A equivalents)", "m", NUM1,
         every("={c}{sh1}+{c}{sh2}+{c}{sh3}", ONON="={c}{sh1}+{c}{sh2}/10"), note="On Class B counted at 1/10 (see notes).")
    line("mcap_list", "Market cap, listing currency", "m", NUM,
         every("={c}{price}*{c}{econ}", UAA="={c}{price}*({c}{sh1}+{c}{sh2})+{c}{price2}*{c}{sh3}"))
    line("mcap", "Market cap, reporting currency", "m", NUM,
         every("={c}{mcap_list}", ONON="={c}{mcap_list}*{x_fx_chf}/{x_fx_usd}"), bold=True,
         note="On: USD converted to CHF at (CHF per EUR) / (USD per EUR), ECB reference rates on Assumptions.")

    section("Balance sheet (latest reported quarter-end, reporting currency)")
    line("cash", "Cash and equivalents", "m", NUM, {"NKE": "={x_cash}+{x_sti}", **{t: peer(t, "cash") for t in peers}},
         note="Nike includes $1.5bn of short-term investments; Lululemon's line includes restricted cash.")
    line("debt_cur", "Debt: current borrowings", "m", NUM, {"NKE": "={x_curltd}", **{t: peer(t, "debt_current") for t in peers}})
    line("debt_nc", "Debt: non-current borrowings", "m", NUM, {"NKE": "={x_ltd}", **{t: peer(t, "debt_noncurrent") for t in peers}})
    line("debt", "Total debt", "m", NUM, every("={c}{debt_cur}+{c}{debt_nc}"))
    ifrs = [t for t, c in zip(tickers, COMPS) if c[2] == "IFRS"]
    line("lease_cur", "Lease liabilities, current (IFRS only)", "m", NUM, {t: peer(t, "lease_current") for t in ifrs})
    line("lease_nc", "Lease liabilities, non-current (IFRS only)", "m", NUM, {t: peer(t, "lease_noncurrent") for t in ifrs})
    line("leases", "Lease liabilities in EV", "m", NUM, {t: "={c}{lease_cur}+{c}{lease_nc}" for t in ifrs},
         note="US GAAP companies: operating leases excluded from EV (lease cost is inside EBITDA).")
    line("nci", "Non-controlling interests", "m", NUM, {"ADS.DE": peer("ADS.DE", "nci")})
    line("ev", "Enterprise value = market cap + debt + leases + NCI - cash", "m", NUM,
         every("={c}{mcap}+{c}{debt}+{c}{leases}+{c}{nci}-{c}{cash}"), bold=True)

    section("LTM income statement (reporting currency)")
    fy_links = {"revenue": "IS.rev", "ebit": "IS.ebit", "da": "IS.da", "net_income": "IS.ni"}
    for item, label in (("revenue", "Revenue"), ("ebit", "EBIT (operating profit)"), ("da", "Depreciation and amortization"),
                        ("net_income", "Net income to shareholders")):
        for period, plabel in (("FY", "last fiscal year"), ("YTD", "current year-to-date"), ("YTD_prior", "prior year-to-date")):
            cells = {t: peer(t, item, period) for t in peers}
            cells["NKE"] = "={x_" + item + "_fy}" if period == "FY" else nke[(item, period.lower())]
            line(f"{item}_{period.lower()}", f"{label}: {plabel}", "m", NUM, cells)
        ltm = every(f"={{c}}{{{item}_fy}}+{{c}}{{{item}_ytd}}-{{c}}{{{item}_ytd_prior}}")
        if item == "da":
            ltm.update({"ADS.DE": "={c}{da_fy}", "PUM.DE": "n/a"})
        line(f"{item}_ltm", f"{label}: LTM", "m", NUM, ltm, bold=True,
             note="adidas: FY2025 D&A (Annual Report) used for LTM, no nine-month figure is disclosed. PUMA does not disclose D&A."
             if item == "da" else "")
    line("ebitda_ltm", "EBITDA: LTM = EBIT + D&A", "m", NUM, every("={c}{ebit_ltm}+{c}{da_ltm}", PUM_DE="n/a"), bold=True)
    line("adj_nb", "adidas cross-check: adjusted net borrowings", "m", NUM, {"ADS.DE": peer("ADS.DE", "adj_net_borrowings")})
    line("adj_ratio", "adidas cross-check: adjusted net borrowings / EBITDA (reported, rounded)", "x", "0.0",
         {"ADS.DE": peer("ADS.DE", "adj_nb_to_ebitda", "LTM")})
    line("adj_ebitda", "adidas cross-check: EBITDA implied by the leverage ratio", "m", NUM,
         {"ADS.DE": "={c}{adj_nb}/{c}{adj_ratio}"}, note="Not used in the multiples; confirms the EBIT + D&A figure above.")

    section("Multiples and operating metrics")
    line("ev_rev", "EV / Revenue (LTM)", "x", '0.00"x"', every("={c}{ev}/{c}{revenue_ltm}"), bold=True)
    line("ev_ebitda", "EV / EBITDA (LTM)", "x", MULT,
         every('=IF(ISNUMBER({c}{ebitda_ltm}),IF({c}{ebitda_ltm}>0,{c}{ev}/{c}{ebitda_ltm},"N/M"),"N/M")'), bold=True)
    line("pe", "P / E (LTM) = market cap / net income", "x", MULT,
         every('=IF({c}{net_income_ltm}>0,{c}{mcap}/{c}{net_income_ltm},"N/M")'), bold=True)
    line("ebit_m", "EBIT margin (LTM)", "%", PCT, every("={c}{ebit_ltm}/{c}{revenue_ltm}"))
    line("ytd_g", "Revenue growth, year-to-date vs. prior year-to-date", "%", PCT, every("={c}{revenue_ytd}/{c}{revenue_ytd_prior}-1"))
    line("ebitda_m", "EBITDA margin (LTM)", "%", PCT, every('=IF(ISNUMBER({c}{ebitda_ltm}),{c}{ebitda_ltm}/{c}{revenue_ltm},"N/M")'))
    line("mid_m", "Mid-cycle EBITDA margin: FY2022-FY2024 average (pre-FY2025)", "%", PCT, {"NKE": "={x_mid_m}"},
         note="Nike only: simple average of FY2022-24 EBITDA / revenue (IS sheet), before the FY2025-26 reset. FY2022 D&A is "
              "depreciation only (old cash-flow layout), so that year is slightly understated. Peers are not normalized.")
    line("mid_ebitda", "Mid-cycle EBITDA = mid-cycle margin x LTM revenue", "m", NUM, {"NKE": "={c}{mid_m}*{c}{revenue_ltm}"})
    line("ev_mid", "EV / mid-cycle EBITDA", "x", MULT, {"NKE": "={c}{ev}/{c}{mid_ebitda}"})
    usd = {"USD": "={c}{%s}", "EUR": "={c}{%s}*{x_fx_usd}", "CHF": "={c}{%s}*{x_fx_usd}/{x_fx_chf}"}
    for key, label in (("mcap", "Memo: market cap, USD"), ("ev", "Memo: enterprise value, USD")):
        line(f"{key}_usd", label, "USD m", NUM, {t: usd[c[3]] % key for t, c in zip(tickers, COMPS)})

    # peer statistics and implied NKE value, below the main table
    n = len(lines)
    stat_rows = [("st_med", "Peer median", "MEDIAN({r})"), ("st_mean", "Peer mean", "AVERAGE({r})"),
                 ("st_p25", "Peer 25th percentile", "QUARTILE({r},1)"), ("st_p75", "Peer 75th percentile", "QUARTILE({r},3)"),
                 ("st_n", "Peers with a meaningful value", "COUNT({r})")]
    mults = [("ev_rev", "EV / Revenue"), ("ev_ebitda", "EV / EBITDA"), ("pe", "P / E")]
    off = {key: i for i, (key, *_rest) in enumerate(lines) if key}
    off["st_head"] = n + 1
    for i, (key, _, _) in enumerate(stat_rows):
        off[key] = n + 2 + i
    off["st_nke"] = n + 2 + len(stat_rows)
    off["imp_head"] = off["st_nke"] + 2
    implied = mults + [("ev_ebitda_mid", "EV / EBITDA applied to Nike mid-cycle EBITDA")]
    stat_col = {"ev_rev": "C", "ev_ebitda": "D", "pe": "E", "ev_ebitda_mid": "D"}
    for i, (key, _) in enumerate(implied):
        off["imp_" + key] = off["imp_head"] + 1 + i
    off["notes"] = off["imp_" + implied[-1][0]] + 2
    nrows = off["notes"] + 1 + len(COMPS_NOTES)

    def render(ws, r0):
        rows = {k: r0 + o for k, o in off.items()}
        X = {"x_price": S("price"), "x_sha": S("shares_a"), "x_shb": S("shares_b"), "x_fx_usd": S("fx_usd_eur"),
             "x_fx_chf": S("fx_chf_eur"), "x_cash": S("q1_cash", "DCF"), "x_sti": S("q1_sti", "DCF"),
             "x_curltd": S("q1_curltd", "DCF"), "x_ltd": S("q1_ltd", "DCF"), "x_dil": S("dil_shares", "DCF"),
             **{f"x_{item}_fy": R(key, 2026) for item, key in fy_links.items()},
             "x_mid_m": "(" + "+".join(f"{R('IS.ebitda', y)}/{R('IS.rev', y)}" for y in (2022, 2023, 2024)) + ")/3"}
        for col_t, col in cols.items():
            h = ws[f"{col}4"]
            h.value, h.font, h.fill, h.alignment = col_t, font("FFFFFF", bold=True), HEAD_FILL, Alignment(horizontal="center")
        for c, text in (("A", "Reporting currency, millions"), ("B", "Unit"), ("J", "Notes")):
            ws[f"{c}4"], ws[f"{c}4"].font, ws[f"{c}4"].fill = text, font("FFFFFF", bold=True), HEAD_FILL
        for i, (key, label, unit, fmt, cells, note, bold) in enumerate(lines):
            r = r0 + i
            if key is None:
                for c in "ABCDEFGHIJ":
                    ws[f"{c}{r}"].fill = SECTION_FILL
                ws[f"A{r}"], ws[f"A{r}"].font = label, font(bold=True)
                continue
            ws[f"A{r}"], ws[f"A{r}"].font = label, font(bold=bold)
            ws[f"B{r}"], ws[f"B{r}"].font = unit, font(GREY, size=9)
            for t, spec in cells.items():
                c = ws[f"{cols[t]}{r}"]
                if isinstance(spec, tuple):
                    c.value, c.font = spec[0], font(BLUE)
                    c.comment = Comment(spec[1], "build_model.py", width=420, height=140)
                elif spec.startswith("="):
                    c.value = spec.format(c=cols[t], **rows, **X)
                    c.font = font(GREEN if "!" in c.value else BLACK, bold=bold)
                else:
                    c.value, c.font = spec, font(bold=bold)
                c.number_format = fmt
                c.alignment = Alignment(horizontal="right")
            ws[f"J{r}"], ws[f"J{r}"].font = note, font(GREY, size=8)

        hr = rows["st_head"]
        ws[f"A{hr}"] = "Peer statistics (excluding Nike; N/M ignored)"
        for j, (_, mlabel) in enumerate(mults):
            ws.cell(hr, 3 + j, mlabel)
        for c in "ABCDE":
            ws[f"{c}{hr}"].font, ws[f"{c}{hr}"].fill = font(bold=True), SECTION_FILL
        for key, label, fn in stat_rows:
            r = rows[key]
            ws[f"A{r}"], ws[f"A{r}"].font = label, font()
            for j, (mkey, _) in enumerate(mults):
                c = ws.cell(r, 3 + j, "=" + fn.format(r=f"$D${rows[mkey]}:$I${rows[mkey]}"))
                c.font, c.number_format = font(), "0" if key == "st_n" else MULT
        r = rows["st_nke"]
        ws[f"A{r}"], ws[f"A{r}"].font = "NIKE (current, for reference)", font(bold=True)
        for j, (mkey, _) in enumerate(mults):
            c = ws.cell(r, 3 + j, f"=$C${rows[mkey]}")
            c.font, c.number_format = font(bold=True), MULT

        hr = rows["imp_head"]
        for col, text in ((1, "Implied NKE value per share (US$): peer multiple x Nike LTM metric"), (3, "25th pct"), (4, "Median"), (5, "75th pct")):
            ws.cell(hr, col, text)
        for c in "ABCDE":
            ws[f"{c}{hr}"].font, ws[f"{c}{hr}"].fill = font(bold=True), SECTION_FILL
        nd = f"($C${rows['debt']}-$C${rows['cash']})"
        metric = {"ev_rev": f"$C${rows['revenue_ltm']}", "ev_ebitda": f"$C${rows['ebitda_ltm']}",
                  "ev_ebitda_mid": f"$C${rows['mid_ebitda']}"}
        for mkey, mlabel in implied:
            r = rows["imp_" + mkey]
            ws[f"A{r}"], ws[f"A{r}"].font = mlabel, font()
            for k, stat in enumerate(("st_p25", "st_med", "st_p75")):
                m = f"{stat_col[mkey]}${rows[stat]}"
                if mkey == "pe":
                    f = f"={m}*$C${rows['net_income_ltm']}/{X['x_dil']}"
                else:
                    f = f"=({m}*{metric[mkey]}-{nd})/{X['x_dil']}"
                c = ws.cell(r, 3 + k, f)
                c.font, c.number_format = font(GREEN), USD2
            ws[f"J{r}"] = "EV multiples: (multiple x Nike LTM metric - Nike net debt) / diluted shares; P/E: multiple x Nike LTM net income / diluted shares."
            ws[f"J{r}"].font = font(GREY, size=8)

        r = rows["notes"]
        for c in "ABCDEFGHIJ":
            ws[f"{c}{r}"].fill = SECTION_FILL
        ws[f"A{r}"], ws[f"A{r}"].font = "Notes: accounting (IFRS vs. US GAAP), currency and data caveats", font(bold=True)
        for i, text in enumerate(COMPS_NOTES):
            c = ws[f"A{r + 1 + i}"]
            c.value, c.font = text, font(size=9)
            ws.merge_cells(f"A{r + 1 + i}:J{r + 1 + i}")
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r + 1 + i].height = 40

    sh.items.append(Block(nrows, render, {k: o for k, o in off.items()}))
    return sh


def build_summary():
    sh = Sheet("Summary", "NIKE, Inc. (NKE) - Valuation summary", years=False, freeze=None,
               subtitle="US$ per share. DCF ranges = +/- one sensitivity step on WACC and terminal growth / exit multiple "
                        "(steps on Assumptions); comps ranges = peer 25th-75th percentile multiples applied to Nike LTM figures.",
               widths={"A": 58, "B": 3, "C": 13, "D": 13, "E": 14, "F": 16, "G": 13, "H": 60})

    def render(ws, r0):
        CUR["sheet"] = "Summary"
        dcf = lambda k: S(k, "DCF")
        comps = lambda k, col: f"Comps!${col}${REG[('Comps', k)]}"
        pg, em = REG[("DCF", "pg_hdr")], REG[("DCF", "em_hdr")]
        info = (("Current share price (close)", f"={S('price')}", USD2), ("Pricing / valuation date", f"={S('price_date')}", "yyyy-mm-dd"),
                ("Operating scenario", f"={S('scenario')}", "@"), ("WACC", f"={dcf('wacc')}", PCT2))
        for i, (label, f, fmt) in enumerate(info):
            ws[f"A{r0 + i}"], ws[f"A{r0 + i}"].font = label, font(bold=i == 0)
            c = ws[f"C{r0 + i}"]
            c.value, c.number_format, c.font = f, fmt, font(GREEN, bold=i == 0)

        h = r0 + len(info) + 1
        for col, text in (("A", "Method"), ("C", "Low"), ("D", "High"), ("E", "Base / median"), ("F", "Base vs. current"),
                          ("G", "Bar length"), ("H", "Basis")):
            c = ws[f"{col}{h}"]
            c.value, c.font, c.fill = text, font("FFFFFF", bold=True), HEAD_FILL
        methods = [
            ("52-week trading range (daily closes)", f"={S('low_52w')}", f"={S('high_52w')}", None,
             "Lowest and highest close, 2025-10-31 to 2026-10-30."),
            ("DCF: perpetuity growth", f"=MIN(DCF!$E${pg + 4},DCF!$G${pg + 2})", f"=MAX(DCF!$E${pg + 4},DCF!$G${pg + 2})",
             f"={dcf('px_g')}", "Low: WACC +1 step, growth -1 step; high: WACC -1 step, growth +1 step (DCF grid)."),
            ("DCF: exit multiple", f"=MIN(DCF!$E${em + 4},DCF!$G${em + 2})", f"=MAX(DCF!$E${em + 4},DCF!$G${em + 2})",
             f"={dcf('px_m')}", "Low: WACC +1 step, multiple -1 step; high: WACC -1 step, multiple +1 step (DCF grid)."),
            ("Comps: EV / EBITDA (LTM)", f"=MIN({comps('imp_ev_ebitda', 'C')},{comps('imp_ev_ebitda', 'E')})",
             f"=MAX({comps('imp_ev_ebitda', 'C')},{comps('imp_ev_ebitda', 'E')})", f"={comps('imp_ev_ebitda', 'D')}",
             "Peer 25th-75th percentile, base = median (4 peers with EBITDA: adidas, Lululemon, Deckers, On)."),
            ("Comps: EV / EBITDA on mid-cycle EBITDA", f"=MIN({comps('imp_ev_ebitda_mid', 'C')},{comps('imp_ev_ebitda_mid', 'E')})",
             f"=MAX({comps('imp_ev_ebitda_mid', 'C')},{comps('imp_ev_ebitda_mid', 'E')})", f"={comps('imp_ev_ebitda_mid', 'D')}",
             "Same peer multiples applied to Nike LTM revenue x FY2022-24 average EBITDA margin (normalized earnings)."),
            ("Comps: P / E (LTM)", f"=MIN({comps('imp_pe', 'C')},{comps('imp_pe', 'E')})",
             f"=MAX({comps('imp_pe', 'C')},{comps('imp_pe', 'E')})", f"={comps('imp_pe', 'D')}",
             "Peer 25th-75th percentile, base = median (profitable peers only)."),
            ("Comps: EV / Revenue (LTM)", f"=MIN({comps('imp_ev_rev', 'C')},{comps('imp_ev_rev', 'E')})",
             f"=MAX({comps('imp_ev_rev', 'C')},{comps('imp_ev_rev', 'E')})", f"={comps('imp_ev_rev', 'D')}",
             "Peer 25th-75th percentile, base = median; ignores margin differences, so weakest of the three."),
        ]
        first = h + 1
        for i, (label, lo, hi, base, basis) in enumerate(methods):
            r = first + i
            ws[f"A{r}"], ws[f"A{r}"].font = label, font(bold=True)
            for col, f, fmt in (("C", lo, USD2), ("D", hi, USD2), ("E", base, USD2),
                                ("F", f"=E{r}/$C${r0}-1" if base else None, PCT), ("G", f"=D{r}-C{r}", USD2)):
                if f:
                    c = ws[f"{col}{r}"]
                    c.value, c.number_format = f, fmt
                    c.font = font(GREEN if "!" in f else BLACK)
            ws[f"H{r}"], ws[f"H{r}"].font = basis, font(GREY, size=8)
        last = first + len(methods) - 1
        r = last + 1
        ws[f"A{r}"], ws[f"A{r}"].font = "Current share price (reference line)", font(italic=True)
        ws[f"E{r}"] = f"=$C${r0}"
        ws[f"E{r}"].number_format, ws[f"E{r}"].font = USD2, font(italic=True)
        ws[f"A{r + 1}"] = ("Column G is the chart helper (high - low). The chart stacks an invisible 'low' bar under the visible range. "
                           "DCF values move with the scenario selector on Assumptions; comps and the 52-week range do not.")
        ws[f"A{r + 1}"].font = font(GREY, italic=True, size=8)

        ch = BarChart()
        ch.type, ch.grouping, ch.overlap, ch.gapWidth = "bar", "stacked", 100, 60
        low = Series(Reference(ws, min_col=3, min_row=first, max_row=last), title="Low")
        rng = Series(Reference(ws, min_col=7, min_row=first, max_row=last), title="Range")
        low.graphicalProperties.noFill = True
        low.graphicalProperties.line.noFill = True
        rng.graphicalProperties.solidFill = "1F3864"
        ch.append(low)
        ch.append(rng)
        ch.set_categories(Reference(ws, min_col=1, min_row=first, max_row=last))
        ch.title = "NKE value per share (US$): football field"
        ch.y_axis.title = "US$ per share"
        ch.x_axis.scaling.orientation = "maxMin"  # first method at the top
        ch.x_axis.delete = False
        ch.y_axis.delete = False
        ch.legend = None
        ch.height, ch.width = 9, 24
        ws.add_chart(ch, f"A{r + 3}")

    sh.items.append(Block(30, render))
    return sh


def build_checks():
    sh = Sheet("Checks", "NIKE, Inc. - Model integrity checks", hist_years=BS_HIST)
    it = sh.items
    sh.section("Reported subtotals (inputs from filings, used only for tie-outs)")
    reported = [
        ("r_gp", "IS", "us-gaap_GrossProfit", "Gross profit"),
        ("r_sga", "IS", "us-gaap_SellingGeneralAndAdministrativeExpense", "Total selling and administrative expense"),
        ("r_pbt", "IS", "us-gaap_IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest", "Income before income taxes"),
        ("r_ni", "IS", "us-gaap_NetIncomeLoss", "Net income"),
        ("r_eps", "IS", "us-gaap_EarningsPerShareDiluted", "Diluted EPS ($)"),
        ("r_tca", "BS", "us-gaap_AssetsCurrent", "Total current assets"),
        ("r_ta", "BS", "us-gaap_Assets", "Total assets"),
        ("r_tcl", "BS", "us-gaap_LiabilitiesCurrent", "Total current liabilities"),
        ("r_eq", "BS", "us-gaap_StockholdersEquity", "Total shareholders' equity"),
        ("r_tle", "BS", "us-gaap_LiabilitiesAndStockholdersEquity", "Total liabilities and shareholders' equity"),
        ("r_cfo", "CF", "us-gaap_NetCashProvidedByUsedInOperatingActivities", "Cash provided by operations"),
        ("r_cfi", "CF", "us-gaap_NetCashProvidedByUsedInInvestingActivities", "Cash used by investing"),
        ("r_cff", "CF", "us-gaap_NetCashProvidedByUsedInFinancingActivities", "Cash used by financing"),
        ("r_end", "CF", "us-gaap_CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents[end]", "Ending cash"),
    ]
    for key, st_name, lid, label in reported:
        h, n = stmt(st_name, lid)
        h = {y: v for y, v in h.items() if y in (BS_HIST if st_name == "BS" else HIST)}
        sh.row(key, f"Reported: {label}", "$" if key == "r_eps" else "$m", hist=h, fmt=USD2 if key == "r_eps" else NUM,
               note=f"{st_name} sheet source filings (row 5 of that sheet); {n}")
    sh.blank()

    sh.section("Checks (absolute differences; status in column B)")
    checks = []

    def check(key, label, years, f, tol=0.5):
        checks.append(key)
        sh.row(key, label, f"tol {tol}", hist=lambda y: f(y) if y in years and y <= 2026 else None,
               fc=(lambda y: f(y)) if any(y in FC for y in years) else None, fmt='#,##0.000;(#,##0.000);"-"')
        TOL[key] = tol

    diff = lambda a, b: (lambda y: f"=ABS({R(a, y)}-{R(b, y)})")
    check("c_bal", "Balance sheet balances (assets = liabilities + equity)", BS_HIST + FC, diff("BS.ta", "BS.tle"))
    check("c_cash", "CF ending cash = BS cash", HIST + FC, diff("CF.end_cash", "BS.cash"))
    check("c_seg", "Segment revenue build = IS revenue", HIST + FC, diff("Segments.total", "IS.rev"))
    check("c_segebit", "Segment EBIT = IS EBIT (history)", HIST, diff("Segments.tot_ebit", "IS.ebit"))
    for key, label, model in (("r_gp", "Gross profit", "IS.gp"), ("r_sga", "Total SG&A", "IS.sga"), ("r_pbt", "Pre-tax income", "IS.pbt"),
                              ("r_ni", "Net income", "IS.ni"), ("r_tca", "Total current assets", "BS.tca"), ("r_ta", "Total assets", "BS.ta"),
                              ("r_tcl", "Total current liabilities", "BS.tcl"), ("r_eq", "Total equity", "BS.eq"),
                              ("r_tle", "Total liabilities and equity", "BS.tle"), ("r_cfo", "Cash from operations", "CF.cfo"),
                              ("r_cfi", "Cash from investing", "CF.cfi"), ("r_cff", "Cash from financing", "CF.cff"),
                              ("r_end", "Ending cash", "CF.end_cash")):
        yrs = BS_HIST if model.startswith("BS") else HIST
        check("c_" + key, f"Model {label} = reported", yrs, diff(model, key))
    check("c_eps", "Model diluted EPS = reported (rounded to cents)", HIST, diff("IS.eps_d", "r_eps"), tol=0.006)
    check("c_negcash", "Forecast cash below zero (amount short)", FC, lambda y: f"=MAX(0,-{R('BS.cash', y)})")
    sh.blank()

    sh.section("Formula error cells by sheet (any Excel error type; must be 0)")
    for name in ("Assumptions", "Segments", "IS", "BS", "CF", "Ratios", "DCF", "Comps", "Summary"):
        it.append(Scalar(f"err_{name}", f"Error cells on {name}", "count",
                         (lambda n=name: f"=SUMPRODUCT(--ISERROR({n}!$A$1:$N$300))"), fmt="0"))
        checks.append(f"err_{name}")
    sh.blank()

    sh.section("Valuation checks (0 = OK, 1 = problem)")
    grid = lambda k: REG[("DCF", k)] + 3  # centre row of a sensitivity grid (header row + 3)
    for key, label, f in (
            ("err_beta", "Beta: model SLOPE matches the Python OLS cross-check (within 0.001)",
             lambda: f"=IF(ABS({S('beta_raw', 'DCF')}-{S('beta_check', 'DCF')})<0.001,0,1)"),
            ("err_tv", "WACC exceeds terminal growth (perpetuity value defined)",
             lambda: f"=IF({S('wacc', 'DCF')}>{S('tg')},0,1)"),
            ("err_grid", "Sensitivity grid centres equal the DCF values per share",
             lambda: f"=IF(AND(ABS(DCF!$F${grid('pg_hdr')}-{S('px_g', 'DCF')})<0.005,"
                     f"ABS(DCF!$F${grid('em_hdr')}-{S('px_m', 'DCF')})<0.005,"
                     f"ABS(DCF!$F${grid('mg_hdr')}-{S('px_g', 'DCF')})<0.005),0,1)"),
            ("err_wacc", "WACC weights sum to 100%",
             lambda: f"=IF(ABS({S('we', 'DCF')}+{S('wd', 'DCF')}-1)<0.000001,0,1)")):
        it.append(Scalar(key, label, "flag", f, fmt="0"))
        checks.append(key)
    sh.blank()
    it.append(Scalar("master", "ALL CHECKS PASS", "flag", None, fmt="@"))  # formula written by finish_checks
    CHECK_KEYS.extend(checks)
    return sh


TOL, CHECK_KEYS = {}, []


# ---------------------------------------------------------------- rendering

def render(wb, sh):
    ws = wb.create_sheet(sh.name)
    CUR["sheet"] = sh.name
    ws.sheet_view.showGridLines = False
    ws["A1"] = sh.title
    ws["A1"].font = font(bold=True, size=14)
    ws["A2"] = sh.subtitle or "USD millions unless stated. Fiscal year ends May 31 (FY2026 = year ended May 31, 2026). A = actual, E = estimate."
    ws["A2"].font = font(GREY, italic=True, size=9)
    ws["A3"] = f'="Scenario: "&{S("scenario")}'
    ws["A3"].font = font(bold=True, color="C00000")
    if sh.years:
        ws["A4"], ws["B4"], ws["N4"] = "USD millions", "Unit", "Source / notes"
        for y in YEARS:
            ws[f"{COL[y]}4"] = f"FY{y}{'A' if y <= 2026 else 'E'}"
        if sh.c_label:
            ws["C4"] = sh.c_label
        for c in "ABCDEFGHIJKLMN":
            cell = ws[f"{c}4"]
            cell.font, cell.fill = font("FFFFFF", bold=True), HEAD_FILL
            cell.alignment = Alignment(horizontal="center" if c not in "AN" else "left")
        ws["A5"] = "Source filing (hover for accession and link)" if sh.sources else ""
        ws["A5"].font = font(GREY, italic=True, size=8)
        for y, acc in sh.sources.items():
            if y in YEARS and (y in sh.hist_years):
                c = ws[f"{COL[y]}5"]
                c.value = FILING_LABEL[acc]
                c.font, c.alignment = font(GREY, size=8), Alignment(horizontal="center")
                c.comment = Comment(f"{acc}\n{FILING_URL[acc]}", "build_model.py")
        for y in FC:
            ws[f"{COL[y]}5"].fill = FC_FILL

    r = FIRST_ROW
    for item in sh.items:
        if item is None:
            r += 1
        elif isinstance(item, Section):
            for c in "ABCDEFGHIJKLMN":
                ws[f"{c}{r}"].fill = SECTION_FILL
            ws[f"A{r}"] = item.text
            ws[f"A{r}"].font = font(bold=True)
            r += 1
        elif isinstance(item, Text):
            ws[f"A{r}"] = item.text
            ws[f"A{r}"].font = font(italic=item.italic, size=9)
            r += 1
        elif isinstance(item, Block):
            item.render(ws, r)
            CUR["sheet"] = sh.name
            r += item.nrows
        elif isinstance(item, Scalar):
            ws[f"A{r}"], ws[f"B{r}"] = item.label, item.unit
            ws[f"A{r}"].font, ws[f"B{r}"].font = font(), font(GREY, size=9)
            v = item.value() if callable(item.value) else item.value
            c = ws[f"C{r}"]
            c.value, c.number_format = v, item.fmt
            is_formula = isinstance(v, str) and v.startswith("=")
            c.font = font(GREEN if is_formula and "!" in v else BLACK if is_formula else BLUE, bold=item.key == "master")
            if item.key_input:
                c.fill = YELLOW
            ws[f"N{r}"] = item.note
            ws[f"N{r}"].font = font(GREY, size=8)
            r += 1
        else:
            render_row(ws, sh, item, r)
            r += 1

    if sh.years:
        ws.column_dimensions["A"].width = 58
        ws.column_dimensions["B"].width = 8
        for y in YEARS:
            ws.column_dimensions[COL[y]].width = 11
        ws.column_dimensions["C"].width = 12 if sh.c_label else 11
        ws.column_dimensions["N"].width = 90
        if not sh.show_2021 and not sh.c_label and sh.name != "Checks":
            ws.column_dimensions["C"].hidden = True
    for col, w in sh.widths.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = sh.freeze
    return ws


def render_row(ws, sh, row, r):
    ws[f"A{r}"] = row.label
    ws[f"A{r}"].font = font(bold=row.bold)
    ws[f"A{r}"].alignment = Alignment(indent=row.indent)
    ws[f"B{r}"] = row.unit
    ws[f"B{r}"].font = font(GREY, size=9)
    hist_years = row.hist_years or sh.hist_years
    for y in YEARS:
        if y in FC:
            src = row.fc
        else:
            src = row.hist if y in hist_years else None
        if src is None:
            v = None
        elif callable(src):
            v = src(y)
        else:
            v = src.get(y)
        c = ws[f"{COL[y]}{r}"]
        if y in FC:
            c.fill = FC_FILL
        if v is None:
            continue
        c.value, c.number_format = v, row.fmt
        is_formula = isinstance(v, str) and v.startswith("=")
        c.font = font(GREEN if is_formula and "!" in v else BLACK if is_formula else BLUE, bold=row.bold)
        if row.total:
            c.border = Border(top=THIN)
    if row.total:
        ws[f"A{r}"].border = Border(top=THIN)
    ws[f"N{r}"] = row.note
    ws[f"N{r}"].font = font(GREY, size=8)


def finish_checks(ws):
    """Status column (B) for every check row, and the master flag."""
    CUR["sheet"] = "Checks"
    flags = []
    for key in CHECK_KEYS:
        r = REG[("Checks", key)]
        if key.startswith("err_"):
            f = f'=IF(C{r}=0,"OK","ERROR")'
        else:
            f = f'=IF(MAX(C{r}:M{r})<{TOL[key]},"OK","ERROR")'
            ws[f"B{r}"].value = None
        target = f"B{r}" if not key.startswith("err_") else f"D{r}"
        ws[target] = f
        ws[target].font = font(bold=True)
        flags.append(target)
        if not key.startswith("err_"):
            ws[f"N{r}"] = f"Tolerance {TOL[key]}"
            ws[f"N{r}"].font = font(GREY, size=8)
    m = REG[("Checks", "master")]
    ws[f"C{m}"] = "=IF(AND(" + ",".join(f'{f}="OK"' for f in flags) + '),"OK","ERROR")'
    ws[f"C{m}"].font = font(bold=True, color="C00000")


def add_scenario_dropdown(ws):
    r = REG[("Assumptions", "scenario")]
    dv = DataValidation(type="list", formula1='"Base,Bull,Bear"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(f"C{r}")


# ---------------------------------------------------------------- static audit

REF = re.compile(r"(?:(\w+)!)?(\$?)([A-N])(\$?)(\d+)")


def relative(formula, col):
    """Formula with every relative column replaced by its offset from `col`, so a copied-across row compares equal."""
    def sub(m):
        sheet, cabs, c, rabs, row = m.groups()
        cc = c if cabs else f"[{ord(c) - ord(col):+d}]"
        return f"{sheet + '!' if sheet else ''}{cabs}{cc}{rabs}{row}"
    return REF.sub(sub, formula)


def audit(wb, skip=()):
    """No hardcoded values in forecast columns (except Assumptions), and one formula pattern per forecast row.
    Free-form sheets (skip) have no forecast columns and are left out."""
    problems = []
    fc_cols = [COL[y] for y in FC]
    for ws in wb.worksheets:
        if ws.title in skip:
            continue
        for r in range(FIRST_ROW, ws.max_row + 1):
            vals = [ws[f"{c}{r}"].value for c in fc_cols]
            if ws.title != "Assumptions":
                for c, v in zip(fc_cols, vals):
                    if v is not None and not (isinstance(v, str) and v.startswith("=")):
                        problems.append(f"{ws.title}!{c}{r}: hardcoded value {v!r} in a forecast column")
            formulas = [v for v in vals if isinstance(v, str) and v.startswith("=")]
            if formulas and len(formulas) != len(fc_cols):
                problems.append(f"{ws.title}!row {r}: formulas in only {len(formulas)} of {len(fc_cols)} forecast columns")
            if len(formulas) == len(fc_cols):
                patterns = {relative(f, c) for f, c in zip(vals, fc_cols)}
                if len(patterns) > 1:
                    problems.append(f"{ws.title}!row {r}: inconsistent formulas across forecast columns: {sorted(patterns)[:2]}")
    return problems


def main():
    sheets = [build_assumptions(), build_segments(), build_is(), build_bs(), build_cf(), build_ratios(),
              build_dcf(), build_comps(), build_summary(), build_checks()]
    for sh in sheets:
        sh.layout()
    wb = Workbook()
    wb.remove(wb.active)
    for sh in sheets:
        ws = render(wb, sh)
        if sh.name == "Assumptions":
            add_scenario_dropdown(ws)
        if sh.name == "Checks":
            finish_checks(ws)
    wb.save(OUT)
    problems = audit(wb, skip=[sh.name for sh in sheets if not sh.years])
    n_formulas = sum(1 for ws in wb.worksheets for row in ws.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith("="))
    print(f"Wrote {OUT}: {len(wb.worksheets)} sheets, {n_formulas} formulas")
    print(f"Static audit: {len(problems)} problem(s)")
    for p in problems:
        print("  ", p)
    return 1 if problems else 0


if __name__ == "__main__":
    raise SystemExit(main())
