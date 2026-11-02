"""
Market data and peer financials for the valuation (Stage 3).

Pulls, for the pricing date 2026-10-30 (last trading day before the valuation date):
    - 10-year US Treasury par yield (treasury.gov daily yield curve CSV)
    - closing prices for Nike and the six peers (Yahoo Finance chart API), with a
      Nasdaq cross-check of Nike's close, and Nike's 52-week closing range
    - EUR/USD and EUR/CHF reference rates (ECB data API)
    - 61 month-end adjusted closes for NKE and the S&P 500 total return index
      (Yahoo), giving 60 monthly returns for the beta regression
    - peer income statement (FY, current YTD, prior YTD) and balance sheet facts
      from the SEC XBRL instances of Lululemon, Deckers, Under Armour and On Holding
and merges data/manual/manual_inputs.csv (adidas and PUMA figures typed from their
IR releases with page references, plus Damodaran ERP and default spreads).

Usage:
    python market_data.py              # download (cached) + build CSVs
    python market_data.py --refresh    # ignore the cache and re-download

Outputs (data/clean/):
    market_inputs.csv   one row per market input: value, unit, date, source
    beta_returns.csv    month-end adjusted closes and monthly returns, NKE vs. S&P 500 TR
    peers.csv           peer facts in local currency, one row per ticker x item x period

Raw downloads are cached, never edited, in data/raw/market/ and data/raw/peers/.
"""
import argparse
import csv
import json
import os
from datetime import date, datetime, timedelta, timezone

import pandas as pd
import requests

from data_pipeline import CLEAN_DIR, RAW_DIR, fetch, parse_instance

PRICE_DATE = date(2026, 10, 30)  # Friday; the valuation date 2026-11-02 is the next trading day
BETA_FIRST_MONTH = "2021-10"  # 61 month-ends -> 60 monthly returns, Nov 2021 - Oct 2026
BETA_LAST_MONTH = "2026-10"
MARKET_DIR = os.path.join(RAW_DIR, "market")
PEER_DIR = os.path.join(RAW_DIR, "peers")
MANUAL_PATH = os.path.join(os.path.dirname(RAW_DIR), "manual", "manual_inputs.csv")

# Non-SEC sites get a generic User-Agent (the SEC one carries a contact email meant for the SEC only).
web = requests.Session()
web.headers.update({"User-Agent": "Mozilla/5.0 (equity research data script)", "Accept": "*/*"})

TICKERS = ["NKE", "ADS.DE", "PUM.DE", "LULU", "DECK", "ONON", "UAA", "UA"]

# Peers that file XBRL with the SEC. Periods are each company's own fiscal calendar.
SEC_PEERS = {
    "LULU": {"cik": "1397187", "currency": "USD",
             "annual": ("0001397187-26-000020", "lulu-20260201_htm.xml"),
             "interim": ("0001397187-26-000127", "lulu-20260802_htm.xml"),
             "FY": ("2025-02-03", "2026-02-01"), "YTD": ("2026-02-02", "2026-08-02"),
             "YTD_prior": ("2025-02-03", "2025-08-03")},
    "DECK": {"cik": "910521", "currency": "USD",
             "annual": ("0001628280-26-037664", "deck-20260331_htm.xml"),
             "interim": ("0001628280-26-068116", "deck-20260930_htm.xml"),
             "FY": ("2025-04-01", "2026-03-31"), "YTD": ("2026-04-01", "2026-09-30"),
             "YTD_prior": ("2025-04-01", "2025-09-30")},
    "UAA": {"cik": "1336917", "currency": "USD",
            "annual": ("0001336917-26-000073", "ua-20260331_htm.xml"),
            "interim": ("0001336917-26-000111", "ua-20260630_htm.xml"),
            "FY": ("2025-04-01", "2026-03-31"), "YTD": ("2026-04-01", "2026-06-30"),
            "YTD_prior": ("2025-04-01", "2025-06-30")},
    "ONON": {"cik": "1858985", "currency": "CHF",
             "annual": ("0001858985-26-000008", "onholdingag-20251231_htm.xml"),
             "interim": ("0001858985-26-000018", "onholdingag-20260630_htm.xml"),
             "FY": ("2025-01-01", "2025-12-31"), "YTD": ("2026-01-01", "2026-06-30"),
             "YTD_prior": ("2025-01-01", "2025-06-30")},
}

# Flow items: the first concept the company reports wins.
FLOW_CONCEPTS = {
    "revenue": ["us-gaap_RevenueFromContractWithCustomerExcludingAssessedTax", "ifrs-full_RevenueFromContractsWithCustomers"],
    "ebit": ["us-gaap_OperatingIncomeLoss", "ifrs-full_ProfitLossFromOperatingActivities"],
    "net_income": ["us-gaap_NetIncomeLoss", "ifrs-full_ProfitLossAttributableToOwnersOfParent"],
    "da": ["us-gaap_DepreciationDepletionAndAmortization", "us-gaap_DepreciationAmortizationAndAccretionNet",
           "ifrs-full_AdjustmentsForDepreciationAndAmortisationExpense"],  # cash flow statement add-back
}

# Balance sheet items per peer, at the interim balance sheet date. An empty list means the
# balance sheet has no such line (recorded as 0 with that note, not silently).
BS_CONCEPTS = {
    "LULU": {"cash": ["us-gaap_CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents"],
             "debt_current": ["us-gaap_ShortTermBorrowings"], "debt_noncurrent": []},
    "DECK": {"cash": ["us-gaap_CashAndCashEquivalentsAtCarryingValue"], "debt_current": [], "debt_noncurrent": []},
    "UAA": {"cash": ["us-gaap_CashAndCashEquivalentsAtCarryingValue"],
            "debt_current": ["us-gaap_LongTermDebtCurrent"], "debt_noncurrent": ["us-gaap_LongTermDebtNoncurrent"]},
    "ONON": {"cash": ["ifrs-full_CashAndCashEquivalents"], "debt_current": [], "debt_noncurrent": [],
             "lease_current": ["ifrs-full_CurrentLeaseLiabilities"],
             "lease_noncurrent": ["ifrs-full_NoncurrentLeaseLiabilities"]},
}

# Share counts: (item, concept, class member or None). Latest date in the interim filing wins.
SHARE_CONCEPTS = {
    "LULU": [("shares", "dei_EntityCommonStockSharesOutstanding", None),
             ("shares_exchangeable", "shop.lululemon.com_ExchangeableStockSharesOutstanding", None)],
    "DECK": [("shares", "dei_EntityCommonStockSharesOutstanding", None)],
    "UAA": [("shares_class_a", "dei_EntityCommonStockSharesOutstanding", "us-gaap:CommonClassAMember"),
            ("shares_class_b", "dei_EntityCommonStockSharesOutstanding", "us-gaap:ConvertibleCommonStockMember"),
            ("shares_class_c", "dei_EntityCommonStockSharesOutstanding", "us-gaap:CommonClassCMember")],
    "ONON": [("shares_class_a", "ifrs-full_NumberOfSharesOutstanding", "onholdingag:ClassASharesMember"),
             ("shares_class_b", "ifrs-full_NumberOfSharesOutstanding", "onholdingag:ClassBVotingSharesMember")],
}


def get(url, name, refresh):
    """Download a non-SEC URL to data/raw/market/<name> (unless cached) and return the path."""
    path = os.path.join(MARKET_DIR, name)
    if os.path.exists(path) and not refresh:
        return path
    os.makedirs(MARKET_DIR, exist_ok=True)
    resp = web.get(url, timeout=60)
    resp.raise_for_status()
    with open(path, "wb") as f:
        f.write(resp.content)
    return path


def epoch(d):
    return int(datetime(d.year, d.month, d.day, tzinfo=timezone.utc).timestamp())


# --- Market data ----------------------------------------------------------------

def treasury_10y(refresh):
    url = ("https://home.treasury.gov/resource-center/data-chart-center/interest-rates/daily-treasury-rates.csv/"
           f"{PRICE_DATE.year}/all?type=daily_treasury_yield_curve&field_tdr_date_value={PRICE_DATE.year}&page&_format=csv")
    df = pd.read_csv(get(url, f"treasury_{PRICE_DATE.year}.csv", refresh))
    row = df[df.Date == PRICE_DATE.strftime("%m/%d/%Y")]
    return round(float(row["10 Yr"].iloc[0]) / 100, 6), url


def yahoo_chart(symbol, first, last, interval, refresh):
    """Return (bars DataFrame with local date, close, adjclose; currency; url)."""
    url = (f"https://query1.finance.yahoo.com/v8/finance/chart/{symbol}"
           f"?period1={epoch(first)}&period2={epoch(last)}&interval={interval}&events=div%2Csplit")
    path = get(url, f"yahoo_{symbol}_{interval}_{first}_{last}.json", refresh)
    with open(path) as f:
        res = json.load(f)["chart"]["result"][0]
    # Bars are stamped at the session open (daily) or local midnight on the 1st (monthly); both fall on
    # the right calendar date in UTC. Adding meta.gmtoffset would not: it is today's offset (EST), not the
    # bar's (EDT in summer), and would push monthly bars back into the previous month.
    bars = pd.DataFrame({
        "date": [datetime.fromtimestamp(t, timezone.utc).date() for t in res["timestamp"]],
        "close": res["indicators"]["quote"][0]["close"],
        "adjclose": res["indicators"]["adjclose"][0]["adjclose"],
    })
    return bars, res["meta"]["currency"], url


def closing_prices(refresh):
    rows = []
    for t in TICKERS:
        bars, ccy, url = yahoo_chart(t, PRICE_DATE - timedelta(days=6), PRICE_DATE + timedelta(days=1), "1d", refresh)
        close = bars.loc[bars.date == PRICE_DATE, "close"]
        if close.empty or pd.isna(close.iloc[0]):
            raise RuntimeError(f"No {PRICE_DATE} close for {t}")  # never fall back to another day silently
        rows.append({"item": f"price_{t}", "value": round(float(close.iloc[0]), 4), "unit": f"{ccy}/share",
                     "as_of": PRICE_DATE, "source": f"Yahoo Finance chart API, daily close (not dividend-adjusted): {url}"})
    return rows


def nke_52_week_range(refresh):
    """Lowest and highest daily close in the year to the pricing date (football field reference bar)."""
    bars, _, url = yahoo_chart("NKE", PRICE_DATE - timedelta(days=366), PRICE_DATE + timedelta(days=1), "1d", refresh)
    year = bars[(bars.date > PRICE_DATE.replace(year=PRICE_DATE.year - 1)) & (bars.date <= PRICE_DATE)].dropna(subset=["close"])
    print(f"52-week range: {len(year)} daily closes from {year.date.min()} to {year.date.max()}")
    source = f"Yahoo Finance chart API, daily closes {year.date.min()} to {year.date.max()} ({len(year)} sessions): {url}"
    return [{"item": f"nke_52w_{k}", "value": round(float(v), 4), "unit": "USD/share", "as_of": PRICE_DATE, "source": source}
            for k, v in (("low", year.close.min()), ("high", year.close.max()))]


def nasdaq_nke_close(refresh):
    url = (f"https://api.nasdaq.com/api/quote/NKE/historical?assetclass=stocks"
           f"&fromdate={PRICE_DATE - timedelta(days=6)}&todate={PRICE_DATE}&limit=10")
    with open(get(url, f"nasdaq_NKE_{PRICE_DATE}.json", refresh)) as f:
        rows = json.load(f)["data"]["tradesTable"]["rows"]
    close = next(r["close"] for r in rows if r["date"] == PRICE_DATE.strftime("%m/%d/%Y"))
    return float(close.replace("$", "")), url


def ecb_rate(currency, refresh):
    """Units of `currency` per 1 EUR (ECB euro foreign exchange reference rate)."""
    url = (f"https://data-api.ecb.europa.eu/service/data/EXR/D.{currency}.EUR.SP00.A"
           f"?startPeriod={PRICE_DATE}&endPeriod={PRICE_DATE}&format=csvdata")
    df = pd.read_csv(get(url, f"ecb_{currency}_EUR_{PRICE_DATE}.csv", refresh))
    return float(df.OBS_VALUE.iloc[0]), url


def beta_returns(refresh):
    first = date.fromisoformat(BETA_FIRST_MONTH + "-01")
    last = PRICE_DATE + timedelta(days=1)
    series = {}
    for symbol, col in [("NKE", "nke"), ("^SP500TR", "sp500tr")]:
        bars, _, url = yahoo_chart(symbol, first, last, "1mo", refresh)
        bars["month"] = bars.date.map(lambda d: d.strftime("%Y-%m"))
        bars = bars[(bars.month >= BETA_FIRST_MONTH) & (bars.month <= BETA_LAST_MONTH)]
        series[col] = bars.set_index("month").adjclose
        series[col + "_url"] = url
    df = pd.DataFrame({"nke_adjclose": series["nke"], "sp500tr_adjclose": series["sp500tr"]})
    print(f"Beta data: {len(series['nke'])} NKE and {len(series['sp500tr'])} index month-ends; "
          f"{len(df)} after aligning on month, {int(df.isna().any(axis=1).sum())} with a missing price")
    df["nke_return"] = df.nke_adjclose.pct_change()
    df["sp500tr_return"] = df.sp500tr_adjclose.pct_change()
    df = df.reset_index().rename(columns={"index": "month"})
    return df, series["nke_url"], series["sp500tr_url"]


# --- Peer financials from SEC XBRL --------------------------------------------------

def peer_instance(cik, acc, name):
    base = f"https://www.sec.gov/Archives/edgar/data/{cik}/{acc.replace('-', '')}"
    return parse_instance(fetch(f"{base}/{name}", os.path.join(PEER_DIR, acc, name))), f"{base}/{name}"


def pick(facts, concepts, start, end):
    """Undimensioned fact for the exact period; returns (value, concept) or (None, None)."""
    for c in concepts:
        same_start = facts.start.isna() if start is None else (facts.start == start)
        hit = facts[(facts.concept == c) & same_start & (facts.end == end) & (facts.dims_key == ())]
        if len(hit):
            return hit.value.iloc[0], c
    return None, None


def sec_peer_rows(ticker, spec):
    annual, annual_url = peer_instance(spec["cik"], *spec["annual"])
    interim, interim_url = peer_instance(spec["cik"], *spec["interim"])
    ccy = spec["currency"]
    rows = []

    def add(item, period, start, end, value, unit, source):
        rows.append({"ticker": ticker, "item": item, "period": period, "start": start, "end": end,
                     "value": value, "unit": unit, "source": source})

    for item, concepts in FLOW_CONCEPTS.items():
        for period in ["FY", "YTD", "YTD_prior"]:
            start, end = (date.fromisoformat(d) for d in spec[period])
            facts, url = (annual, annual_url) if period == "FY" else (interim, interim_url)
            value, concept = pick(facts, concepts, start, end)
            if value is None:
                raise RuntimeError(f"{ticker} {item} {period}: none of {concepts} reported")
            add(item, period, start, end, value / 1e6, f"{ccy} m", f"{url} | {concept}")

    bs_date = date.fromisoformat(spec["YTD"][1])
    for item, concepts in BS_CONCEPTS[ticker].items():
        value, concept = pick(interim, concepts, None, bs_date) if concepts else (None, None)
        if concepts and value is None:
            raise RuntimeError(f"{ticker} {item}: none of {concepts} reported at {bs_date}")
        if value is None:
            add(item, "BS", None, bs_date, 0.0, f"{ccy} m", f"{interim_url} | no such line on the balance sheet")
        else:
            add(item, "BS", None, bs_date, value / 1e6, f"{ccy} m", f"{interim_url} | {concept}")

    for item, concept, member in SHARE_CONCEPTS[ticker]:
        hits = interim[(interim.concept == concept) & interim.start.isna()]
        hits = hits[hits.dims.map(lambda d: (member in d.values()) if member else not d)]
        latest = hits.sort_values("end").iloc[-1]
        add(item, "BS", None, latest.end, latest.value / 1e6, "shares m",
            f"{interim_url} | {concept}" + (f" [{member}]" if member else ""))
    return rows


# --- Main -------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--refresh", action="store_true")
    refresh = parser.parse_args().refresh

    manual = pd.read_csv(MANUAL_PATH)
    print(f"Manual inputs: {len(manual)} rows, {int(manual.value.isna().sum())} missing values")

    # Market inputs
    rows = closing_prices(refresh)
    nke_nasdaq, nasdaq_url = nasdaq_nke_close(refresh)
    rows.append({"item": "price_NKE_nasdaq_check", "value": nke_nasdaq, "unit": "USD/share", "as_of": PRICE_DATE,
                 "source": f"Nasdaq historical quotes API: {nasdaq_url}"})
    rows += nke_52_week_range(refresh)
    rf, rf_url = treasury_10y(refresh)
    rows.append({"item": "rf_10y", "value": rf, "unit": "fraction", "as_of": PRICE_DATE,
                 "source": f"US Treasury daily par yield curve, 10 Yr column: {rf_url}"})
    for ccy in ["USD", "CHF"]:
        rate, url = ecb_rate(ccy, refresh)
        rows.append({"item": f"fx_{ccy.lower()}_per_eur", "value": rate, "unit": f"{ccy} per EUR", "as_of": PRICE_DATE,
                     "source": f"ECB euro foreign exchange reference rate: {url}"})
    market_manual = manual[manual.ticker == "MARKET"]
    for r in market_manual.itertuples():
        rows.append({"item": r.item, "value": r.value, "unit": r.unit, "as_of": r.end, "source": r.source})
    market = pd.DataFrame(rows)

    nke_yahoo = market.loc[market.item == "price_NKE", "value"].iloc[0]
    if abs(nke_yahoo - nke_nasdaq) > 0.005:
        raise RuntimeError(f"NKE close differs: Yahoo {nke_yahoo} vs Nasdaq {nke_nasdaq}")

    # Beta returns
    beta, nke_url, idx_url = beta_returns(refresh)
    returns = beta.dropna(subset=["nke_return", "sp500tr_return"])
    raw_beta = returns.nke_return.cov(returns.sp500tr_return) / returns.sp500tr_return.var()
    market = pd.concat([market, pd.DataFrame([{
        "item": "beta_raw_check", "value": round(raw_beta, 4), "unit": "x", "as_of": PRICE_DATE,
        "source": f"Python OLS on beta_returns.csv (cross-check of the model's SLOPE); monthly adjusted closes, "
                  f"Yahoo Finance: NKE {nke_url} ; S&P 500 TR {idx_url}"}])])

    # Peers
    peers = [r for t, spec in SEC_PEERS.items() for r in sec_peer_rows(t, spec)]
    peers = pd.concat([pd.DataFrame(peers), manual[manual.ticker != "MARKET"]], ignore_index=True)

    os.makedirs(CLEAN_DIR, exist_ok=True)
    market.to_csv(os.path.join(CLEAN_DIR, "market_inputs.csv"), index=False, encoding="utf-8-sig")
    beta.to_csv(os.path.join(CLEAN_DIR, "beta_returns.csv"), index=False, encoding="utf-8-sig")
    peers.to_csv(os.path.join(CLEAN_DIR, "peers.csv"), index=False, encoding="utf-8-sig", quoting=csv.QUOTE_MINIMAL)

    print(f"market_inputs.csv: {len(market)} rows")
    for r in market.itertuples():
        print(f"  {r.item:24s} {'' if pd.isna(r.value) else r.value:>12} {r.unit:14s} {r.as_of}")
    print(f"beta_returns.csv: {len(beta)} month-ends, {len(returns)} monthly returns "
          f"({returns.month.iloc[0]} to {returns.month.iloc[-1]}); raw OLS beta {raw_beta:.3f}")
    print(f"peers.csv: {len(peers)} rows ({len(peers) - len(manual[manual.ticker != 'MARKET'])} from SEC XBRL, "
          f"{len(manual[manual.ticker != 'MARKET'])} manual), {int(peers.value.isna().sum())} missing values")
    print(peers.pivot_table(index=["ticker"], columns="item", values="value", aggfunc="count").fillna(0).astype(int).to_string())


if __name__ == "__main__":
    main()
