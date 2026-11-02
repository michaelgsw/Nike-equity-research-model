# NIKE, Inc. (NYSE: NKE): equity research model and memo

An entry-level equity research project on Nike, built from public filings. It has three parts:

- a three-statement operating model with Base, Bull and Bear scenarios;
- a DCF and a trading-comps valuation;
- a two-page investment memo.

Every historical number comes from SEC EDGAR XBRL or a cited company document. Every forecast cell is an Excel formula.

- **Pricing date:** 2026-10-30 close.
- **Valuation date:** 2026-10-30.
- **Latest filing used:** 10-Q for Q1 FY2027 (quarter ended 2026-08-31).
- **Fiscal years** end May 31, so FY2026 = year ended May 31, 2026.
- **Units:** US$ millions unless stated.

*Portfolio exercise, not investment advice.*

## Files

| File | What it is |
|---|---|
| `nike_model.xlsx` | The model: Assumptions, Segments, IS, BS, CF, Ratios, DCF, Comps, Summary, Checks |
| `memo.md` / `memo.pdf` | Investment memo (two pages) |
| `memo_football_field.svg` | Valuation chart used in the memo, drawn from the Summary sheet |
| `data_pipeline.py` | Downloads Nike's 10-K (FY2022–FY2026) and 10-Q (Q1 FY2027) XBRL from EDGAR → `data/clean/` |
| `market_data.py` | Downloads market data and peer filings, and merges the manual peer inputs → `data/clean/` |
| `build_model.py` | Writes `nike_model.xlsx` from `data/clean/` and runs a static formula audit |
| `export_memo.py` | Draws the chart and prints `memo.md` to `memo.pdf` with headless Edge/Chrome |
| `data/manual/manual_inputs.csv` | Figures typed by hand from PDFs, each with a source and page: adidas, PUMA, the Damodaran ERP, default spreads |
| `data/clean/` | Tidy CSVs the model reads; each row carries its source (filing, XBRL tag or URL) |
| `data/raw/` | Download cache, never edited. Not in git: the scripts re-create it |

## How to run

Requirements:

- Python 3.11+ with `pandas`, `openpyxl` and `requests`.
- Excel or LibreOffice, to calculate the workbook.
- Microsoft Edge or Google Chrome, for the PDF.

The SEC requires a User-Agent that names you and gives an email. Set it once per shell:

```powershell
$env:SEC_USER_AGENT = "Your Name your@email.com"
```

Then run the scripts in order:

```bash
python data_pipeline.py
```

```bash
python market_data.py
```

```bash
python build_model.py
```

Next, calculate the workbook. openpyxl writes formulas without values, so open `nike_model.xlsx` in Excel and save it (or recalculate it with LibreOffice). Then:

```bash
python export_memo.py
```

Downloads are cached in `data/raw/`; pass `--refresh` to re-download. Without `SEC_USER_AGENT`, the pipeline still runs from the cache but refuses to call the SEC.

## The model

| Sheet | Contents |
|---|---|
| Assumptions | Scenario selector (Base/Bull/Bear dropdown) and drivers by year: segment × channel growth, underlying gross margin and tariff drag, opex ratios, working-capital day targets, capex, dividends and buybacks. Also WACC inputs, terminal value and sensitivity steps. Each assumption has its rationale in the notes column |
| Segments | Revenue by geography × channel (wholesale / NIKE Direct), Converse and Global Brand Divisions; segment EBIT. History ties to the IS |
| IS, BS, CF | FY2021A/22A–FY2026A as reported, then FY2027E–FY2031E. The balance sheet balances through cash; Checks flags any year where forecast cash would go negative |
| Ratios | Margins, growth, return ratios, and working-capital days (DSO, DIO, DPO) |
| DCF | WACC (CAPM: 60-month regression beta with a Blume adjustment, Treasury Rf, Damodaran ERP) and unlevered FCF. Mid-period discounting with an FY2027 stub. Terminal value two ways: Gordon growth (g = 2.25%) and exit multiple (10.0x EBITDA), with each method's implied counterpart. Three sensitivity grids: WACC × g, WACC × exit multiple, and FY2030E EBIT margin × WACC |
| Comps | adidas, PUMA, Lululemon, Deckers, On, Under Armour: LTM EV/Revenue, EV/EBITDA and P/E, peer quartiles, implied Nike value, and an EV/EBITDA on Nike's mid-cycle (FY2022–24 average) EBITDA |
| Summary | Football field (52-week range, two DCF methods, four comps methods) and a scenario table |
| Checks | Tie-outs to reported subtotals, BS balance, cash roll-forward, segment totals, error-cell counts per sheet, and valuation sanity checks. The top cell must read ALL CHECKS PASS |

Conventions:

- **Blue font:** a hardcoded input, either a historical figure from a filing or an assumption.
- **Black font:** a formula.
- **Green font:** a cross-sheet link, where it is highlighted.
- **Yellow fill:** a key assumption.
- **Sources:**
  - Each historical row names its XBRL tag in the notes column.
  - Each historical column names its source filing in row 5.
  - Hand-typed inputs cite a document and page.

`build_model.py` audits the workbook before saving. It fails on two things:

- any hardcoded number in a forecast column outside Assumptions;
- a row whose forecast formulas do not share one pattern.

## Data sources

- **Nike:** 10-K FY2022–FY2026 and 10-Q Q1 FY2027, from SEC EDGAR (XBRL instance documents, `submissions` API).
- **Peers:**
  - Lululemon, Deckers, Under Armour and On Holding: XBRL from their SEC filings.
  - adidas and PUMA: IR releases and the adidas Annual Report 2025, typed into `manual_inputs.csv` with page references.
- **Market:**
  - U.S. Treasury daily par yield curve (10-year).
  - Yahoo Finance chart API: closes, and month-end adjusted closes for NKE and the S&P 500 TR. Nasdaq is used to cross-check Nike's close.
  - ECB data API: EUR/USD and EUR/CHF.
  - Damodaran Online: implied ERP and ratings-based default spreads.

## Weak spots

These are the places where the model is thinnest, roughly in order of how much they move the answer.

1. **Most of the value is in the terminal value and margin recovery.**
   - Terminal value is 76–77% of EV.
   - The Base case assumes EBIT margin recovers from 8.3% (FY2026) to 12.1% (FY2031E). Each 1pp of FY2030E margin is worth about $3.4 per share.
   - The margin path is my assumption, informed by history and management commentary. No consensus estimates were used.
2. **The comps are thin.**
   - Six peers, with different business models (adidas and PUMA are the closest).
   - Two are N/M on EBITDA or earnings (PUMA does not disclose D&A; Under Armour has negative LTM EBITDA). The medians therefore rest on four or five companies.
   - LTM periods are not aligned: peer quarter-ends range from June to September 2026.
   - adidas LTM D&A uses its FY2025 figure as a proxy, because no nine-month D&A is disclosed.
   - The IFRS vs. US GAAP lease difference is not adjusted.
   - PUMA's share count is a proxy.
   - Multiples are trailing only; no forward consensus multiples.
3. **WACC inputs are single-source.**
   - Beta comes from one 60-month window against one index.
   - The ERP comes from one source (Damodaran's implied premium).
   - Debt is at book value, and leases are excluded from both debt and FCF.
   - One WACC is used for all three scenarios.
4. **Timing mismatches.**
   - Net debt is as of 2026-08-31 (the latest balance sheet), while the valuation date is 2026-10-30. This is disclosed in the model.
   - The FY2027 stub assumes cash flow arrives evenly through the year, but Nike's cash flow is seasonal.
   - The $684m IEEPA tariff refund is treated as already collected in Q1 and removed from FY2027E cash flow.
5. **Simplifications.**
   - A flat 20% forecast tax rate.
   - Stock-based compensation is treated as a cash-like cost (not added back), which is conservative.
   - Working capital moves to target days (DIO 98, DSO 37) by FY2028E; the targets are judgement calls.
   - The FY2027 working-capital release is mostly receivables (the IEEPA refund), not inventory.
6. **Mid-cycle margin.**
   - It uses Nike's own FY2022–24 average EBITDA margin (14.65%).
   - FY2022 D&A is depreciation only, because the older cash-flow layout did not report amortization separately. That understates FY2022 slightly.
   - Peers are not normalized the same way.
7. **Data-handling risks.**
   - adidas and PUMA figures were transcribed by hand from PDFs.
   - Yahoo Finance and Nasdaq price endpoints are unofficial and could change or break.
   - Restated prior-year figures are taken from the most recent filing that reports them (logged in `restatements.csv`).
