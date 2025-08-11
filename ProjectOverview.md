## StockTrades Wheel Strategy Dashboard — Project Overview

### Purpose
Generate an Excel dashboard from brokerage transactions to analyze a wheel options strategy. The script reads a CSV export, cleans and enriches the data, calculates YTD metrics, and produces charts and tables in a formatted Excel workbook.

### Repository Layout
- `wheel_strategy_dashboard.py`: Main script that loads data, derives fields, builds summaries, generates charts, and writes the Excel workbook.
- `Designated_Bene_Individual_XXX885_Transactions_20250811-030603.csv`: Sample input CSV used by default in the script.
- `claude_code_prompt_wheel_dashboard.md`: Design/prompt document describing the intended behavior and deliverables for the script.

### Requirements
- Python 3.13+ (script header references 3.13.2)
- macOS is referenced, but the script is cross‑platform if paths are adjusted
- Packages: `pandas`, `matplotlib`, `xlsxwriter`

Install dependencies:

```bash
pip install pandas matplotlib xlsxwriter
```

### Input Data
The script expects a CSV with headers similar to the sample file:

- `Date`
- `Action`
- `Symbol`
- `Description`
- `Quantity`
- `Price`
- `Fees & Comm`
- `Amount`

Notes:
- Headers are normalized internally by replacing spaces with underscores (e.g., `Fees & Comm` → `Fees_&_Comm`).
- `Date` values may contain text like "08/04/2025 as of 08/01/2025". The first MM/DD/YYYY is extracted to derive `TradeDate`.
- Currency/number fields are cleaned by stripping `$` and commas and converting to numbers.

### What the Script Produces
An Excel file named `wheel_all_transactions_with_dashboard_and_ytd_summary_reverted.xlsx` with these sheets:

- Transactions: Full cleaned ledger plus derived columns (`TradeDate`, `Ticker`, `IsOption`, `Premium_Net`, `CashFlow`, `Category`, ...).
- Summary: All‑time totals (total option premium, cash flow by category, top tickers by premium).
- YTD Summary: Year‑to‑date metrics (options premium, net stock P/L, interest, fees, net P/L including interest and fees).
- Dashboard: Monthly premium and cumulative premium (YTD) table, plus bar and line charts.

### How It Works (High‑Level)
1. Load and clean CSV (`load_and_clean`): normalize headers, parse `TradeDate`, coerce numeric columns.
2. Derive fields (`derive_fields`): detect options, infer `Ticker`, compute `Premium_Net`, `CashFlow`, and `Category`.
3. Build summaries (`build_summary_tables`): all‑time and YTD rollups.
4. Build dashboard data (`build_dashboard_data`): monthly YTD premiums and charts.
5. Write Excel (`write_excel`): formatted sheets and embedded chart image.

### Configuration
Edit these constants near the top of `wheel_strategy_dashboard.py`:

- `INPUT_CSV`: Full path to your brokerage CSV
- `OUTPUT_FOLDER`: Destination directory for the Excel file
- `ANALYSIS_YEAR`: Year used for YTD calculations
- `USE_MONTH_LABELS`: `True` for "Jul 2025" style x‑axis labels; `False` for date strings

The script currently sets:

- `INPUT_CSV`: `/Users/shawnsandy/Code Repos/StockTrades/Designated_Bene_Individual_XXX885_Transactions_20250811-030603.csv`
- `OUTPUT_FOLDER`: `/Users/shawnsandy/Code Repos/StockTrades/`
- `ANALYSIS_YEAR`: `2025`
- `USE_MONTH_LABELS`: `True`

### Usage
Run from the project folder (adjust the Python executable as needed):

```bash
python wheel_strategy_dashboard.py
```

On completion, the script prints rows processed, YTD rows, YTD premium total, YTD net P/L, and the output file path.

### Troubleshooting
- If the script exits with “Input file not found,” update `INPUT_CSV` to your actual CSV path.
- If date parsing fails or `TradeDate` is empty, ensure the `Date` column has an MM/DD/YYYY embedded; the first occurrence is used.
- If numeric columns are missing (e.g., `Fees & Comm`), the script creates them with zeros and continues.
- If YTD has no rows for `ANALYSIS_YEAR`, the dashboard section is skipped but other sheets are still written.

### Future Improvements
- Add CLI arguments or interactive prompts (per the design document) instead of hard‑coded constants.
- Validate input schema and report missing/extra columns explicitly.
- Optional: Save charts as native Excel chart objects in addition to the embedded image.

