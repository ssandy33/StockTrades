# Prompt for Claude Code: Wheel Strategy Transactions Dashboard Script

You are to generate a single, runnable Python script that I can execute locally. The script must:

## First, ask me these questions (one by one) and wait for my answers before writing code:
1. My OS (Windows/macOS/Linux).  
2. My Python version (e.g., 3.9/3.10/3.11).  
3. The full path to the original CSV export of my brokerage transactions (the input).  
4. The desired output folder (default to the same folder as the input if I press Enter).  
5. The analysis year for YTD calculations (default to the current year if I press Enter).  
6. Whether to use **month-year labels** on charts (Y/N, default Y).

## What the script must do:
- Use **pandas**, **matplotlib**, and **xlsxwriter** (install with pip if missing).  
- Read the CSV and robustly clean it:
  - Normalize headers: strip spaces and replace spaces with underscores.  
  - Parse a **TradeDate** from the “Date” column by extracting the **first** MM/DD/YYYY in strings like `08/04/2025 as of 08/01/2025`.  
  - Convert numeric columns to real numbers (not text): at minimum `Price`, `Fees_&_Comm`, and `Amount`; remove `$` and commas before conversion.  
- Derive helper fields:
  - `IsOption`: True if Symbol or Description indicates option (contains “ CALL/PUT ”, “ C ”, “ P ”, or “OPTION”).  
  - `Ticker`: From Symbol (before any spaces); fallback to an all‑caps token in Description.  
  - `Premium_Net`: `Amount` if `IsOption`, else 0.  
  - `CashFlow`: `Amount` (signed).  
  - `Category`: Use `Action` if present; else classify row as:
    - If `IsOption`: “Options - STO/BTC/Assigned/Expired/Other” by keywords in `Action` or `Description` (sell to open, buy to close, assigned, expired).  
    - Else if buy/sell keywords present → “Stock - Buy/Sell”; else “Other”.  
- Sort by `TradeDate`, keep a stable order.

## Sheets to write to the output Excel file (`wheel_all_transactions_with_dashboard_and_ytd_summary_reverted.xlsx`):
1. **Transactions** – full cleaned ledger with all derived columns (`TradeDate`, `Ticker`, `IsOption`, `Premium_Net`, `CashFlow`, `Category`, etc.). Ensure all numeric columns are stored as real numbers (no “number stored as text” warnings).  
2. **Summary** – all‑time roll‑ups:  
   - “Total Premium (options)” = sum of `Premium_Net`.  
   - Category breakdown: sum of `CashFlow` by `Category`.  
   - Premium by ticker: sum of `Premium_Net` by `Ticker`.  
3. **YTD Summary** – for the chosen year only:  
   - **YTD Options Premium** = sum of `Premium_Net`.  
   - **YTD Net Stock Trade P/L** = (cash from “Stock - Sell”) + (cash from “Stock - Buy”).  
   - **YTD Net P/L (Premiums + Stock)** = prior two lines combined.  
   - **YTD Interest** = cash from “Interest” rows (if present).  
   - **YTD Fees** = sum of `Fees_&_Comm`.  
   - **YTD Net P/L (Incl Interest & Fees)** = add interest and subtract fees.  
4. **Dashboard** – tables plus two visuals:  
   - Table: **Monthly Premium (YTD)** = sum of `Premium_Net` grouped by month; and **Cumulative_Premium** (running total).  
   - Chart A: **Monthly Premium Income (YTD)** as a **bar** chart (one bar per month).  
   - Chart B: **Cumulative Premium Income (YTD)** as a **line** chart.  
   - Use month‑year labels (e.g., “Jul 2025”) if I answered Y; otherwise use first‑of‑month dates on the x‑axis.  
   - Do **not** use seaborn; use matplotlib.  

## Formatting & quality requirements:
- Use `xlsxwriter` to write the workbook. Apply numeric formats (currency with 2 decimals) to currency columns, and general number formats to quantities/prices so Excel will not show “Number Stored as Text.”  
- Auto‑fit columns where reasonable.  
- The script should print a short summary at the end: rows processed, YTD premium total, YTD net P/L, and the path to the saved Excel file.  
- Include a `requirements.txt` snippet in a comment and a quick usage note in the script header.  
- Wrap the main logic in `if __name__ == "__main__": main()` and use clear functions (e.g., `load_and_clean()`, `derive_fields()`, `build_ytd_tables()`, `write_excel()`).  
- Handle common errors gracefully: missing columns, unreadable dates, empty YTD set, etc., with friendly messages and exit codes.

## Deliverables:
- Output a **single .py file** in your response, ready to run as‑is.  
- At the very top, echo back my answers (Python version, OS, input path, output folder, year, month‑label choice) as comments so I can verify.

When you’re ready, start by asking me the six setup questions above. After I answer, generate the script.
