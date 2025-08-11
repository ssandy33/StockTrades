## Wheel Strategy — Approach, Insights, Questions, and Recommendations

### What this strategy is
The wheel is an income and accumulation strategy:
- Sell cash‑secured puts on stocks/ETFs you’re willing to own at a discount. If assigned, you buy 100 shares per contract at strike.
- Once assigned, sell covered calls against the shares. If called away, you exit at strike for a gain plus the collected premiums, and you can repeat the wheel.

### Objectives
- Reliable premium income with controlled downside risk
- Accumulate quality shares at target basis levels
- Systematic, rules‑based entries and exits to reduce discretion

### Context from your current data
- Underlyings: mix of large caps/ETFs (`T`, `INTC`, `VZ`, `HPE`, `DVN`, `USO`, `XLE`).
- YTD option premium (cash basis): about $342.
- Cash-based net P/L is negative because stock purchases/assignments are counted as cash outflows before eventual exits.
- Strikes are often close to spot, which increases assignment likelihood and short‑term drawdown risk.

### What success should look like
- Per‑trade ROI: 0.5–1.5% of collateral for 7–14 DTE; average annualized yields in the 15–35% range.
- Assignment/expiration mix consistent with intent: 20–40% assignments can be healthy for accumulation; most others expire worthless or are closed early for a profit.
- Controlled concentration: per‑ticker collateral and portfolio heat within limits.

### KPIs to track (now available in `WheelStrategyOverview.xlsx`)
- Premium income: by month, YTD, and cumulative.
- Per‑trade metrics (YTD): ROI, annualized yield, DTE, outcome (Expired/Assigned), conservative score.
- Rates: assignment rate, expiration rate, average DTE, average ROI/annualized yield.
- Collateral usage: per trade and total (to monitor cash at risk).

### Questions to answer (to refine your rules)
- Entry selection:
  - What is the target delta/OTM% when selling puts and calls? Any IV rank filter?
  - Preferred DTE (7–10 vs 14–21) and minimum liquidity (spread/volume)?
  - Do you avoid earnings/ex‑dividend weeks on covered calls?
- Risk sizing:
  - Max portfolio heat (total potential assignment notional vs cash)?
  - Max per‑ticker collateral and max concurrent positions?
- Management and exits:
  - Close at 50–75% of max profit? Roll rules when tested (roll down/out for credit)?
  - Covered calls ITM near expiration: roll up/out for credit or allow assignment?
  - When do you switch from accumulation to distribution (harvesting CCs vs holding)?
- Review cadence:
  - Weekly review checklist and thresholds for tightening/loosening risk.

### Recommendations
- Strike selection and DTE
  - Prefer lower‑delta entries for a conservative posture: target ~0.15–0.25 delta for puts and calls.
  - Use 7–14 DTE for frequent cycles; consider 21–30 DTE if liquidity is strong and headroom is needed.
  - Avoid selling calls the week of ex‑dividend if assignment would forfeit the dividend unintentionally.
- Ticker universe and diversification
  - Favor liquid, megacap, boring businesses and broad ETFs; be cautious with higher‑volatility products like `USO`.
  - Cap single‑ticker collateral (e.g., ≤20–25% of total) and sector concentration (≤40–50%).
- Management rules
  - Take profits early: buy back at 50–75% of max gain; redeploy into a new cycle.
  - When tested on short puts, roll down and/or out only for net credit with improved break‑even.
  - For covered calls that are ITM, roll up and out for credit if you want to keep shares; otherwise allow assignment.
- Measurement and review
  - Distinguish realized vs unrealized P/L. Use the new `WheelStrategyOverview` and `Options Trades (YTD)` sheets.
  - Track collateral at risk, assignment/expiration rates, and per‑trade annualized yield.
  - Maintain a weekly review: close profitable trades, manage tested positions, stagger expirations.

### Weekly operating checklist
- Screen candidates for liquidity, IV, trend/levels; exclude earnings‑week risks.
- Allocate positions to stay within collateral and concentration limits.
- Enter puts at target delta/OTM% and DTE; place GTC profit targets (e.g., 50–70%).
- Manage: roll tested puts for credit; roll CCs if you want to retain shares; otherwise accept assignment.
- Update logs, review KPIs in `WheelStrategyOverview.xlsx`, and adjust sizing if heat is high.

### Risk notes
- Assignments are expected; ensure cash reserves for potential ownership.
- Gap risk around earnings and macro events can exceed expected OTM cushions.
- Liquidity matters: wide spreads can negate small premiums.

### Next steps
- Use `WheelStrategyOverview.xlsx` after each run to review KPI thresholds.
- If desired, extend automation to include delta/IV data and earnings calendars for stricter filters.

