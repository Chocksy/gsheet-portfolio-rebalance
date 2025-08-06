# Portfolio Tracker & Rebalancer

A Google Sheets + Apps Script project that tracks an equity portfolio, suggests equal-weight trades, and keeps a live dashboard of performance metrics.

---

## 1  Features

* **Custom menu** → `📈 Portfolio Tools › ▶️ Calculate & Rebalance Portfolio` runs the script.
* **Automatic sheet provisioning** – missing sheets & headers are created on first run.
* **Equal-weight logic** – target value per ticker = *Target Allocation* / *number of tickers*.
* **1-share safeguard** – every listed ticker holds at least one share even if price > allocation/stock.
* **Trade history** – every BUY/SELL is appended to `TradeHistory`.
* **Live dashboard** – current prices, P&L, allocation/stock, YTD stats—updated via formulas.
* **Debug logging** – review step-by-step output in Apps Script **Executions › Logs**.

---

## 2  Required Sheets

| Sheet | Purpose | User edits? |
|-------|---------|-------------|
| `TradeList` | Desired tickers list. Column A = Ticker, Column B holds `=GOOGLEFINANCE()` price formulas. | ✅ Yes (add/remove tickers) |
| `Holdings` | Actual share counts & cost basis (managed by script). | 🚫 |
| `AllocationHistory` | Log of total capital additions. Last row = current *Target Allocation*. | ✅ |
| `TradeHistory` | Append-only trade log (managed by script). | 🚫 |
| `Dashboard` | Live performance view (formulas, auto-generated). | 🚫 |

The script will create any missing sheet with correct headers.

---

## 3  Installation / First-time Setup

1. **Copy** the entire folder into a Google Drive location or open the `Code.gs` file in the Apps Script editor attached to your spreadsheet.
2. In Google Sheets **Extensions › Apps Script** paste the contents of `Code.gs`.
3. Save and **run `runRebalance` once** from the Apps Script editor to grant permissions and bootstrap all sheets.
4. Fill in:
   * `AllocationHistory` – add a row with today’s date & your total investable cash.
   * `TradeList` – list tickers in Column A. Column B will auto-populate prices after the first run (or add `=GOOGLEFINANCE(A2,"price")` yourself).
5. Back in the spreadsheet reload → you’ll see **📈 Portfolio Tools** menu.
6. Use **▶️ Calculate & Rebalance Portfolio** any time you change `TradeList` or add capital.

---

## 4  How Rebalancing Works

1. Gather data
   * Desired tickers (`TradeList`)
   * Current holdings (`Holdings`)
   * Latest *Target Allocation* (`AllocationHistory` last row)
2. Price lookup
   * Uses price in `TradeList!B` if numeric; otherwise sets a temporary `GOOGLEFINANCE` formula.
3. Calculates target shares per ticker.
4. Ensures minimum 1 share per ticker.
5. Builds BUY/SELL list, records trades in `TradeHistory`, updates `Holdings`.
6. Refreshes `Dashboard` and shows summary dialog.

---

## 5  Dashboard Reference

* **Columns A-G** are driven by ARRAYFORMULAS—only the first cell shows the formula; downstream cells update automatically.
* **Metric block (J-K)**
  * Target Allocation – last `AllocationHistory` value
  * Portfolio Value – sum of live Market Value column
  * Cash Position – Target Allocation − cost basis
  * Unrealized P&L – difference between market value and cost basis
  * Total P&L – portfolio value minus **first** AllocationHistory entry
  * Starting Capital – first Allocation entry **of current year**
  * YTD P&L – portfolio value minus starting capital
  * YTD P&L % – YTD P&L ÷ starting capital
  * Allocation / Stock – Target Allocation ÷ number of tickers

All metrics refresh automatically as prices or sheets change.

---

## 6  Debugging & Logs

* Set `const DEBUG = true;` (default) at top of `Code.gs` to enable verbose `Logger.log()` output.
* Run `runRebalance` from the Apps Script editor or use the menu; open **Executions › View logs** for details.

---

## 7  Customization Notes

* Adjust minimum-share rule inside `rebalancePortfolio()` if a different fallback is desired.
* Change formatting/column order in `updateDashboard()` to suit your style.
* The price cache is 5 minutes (`CacheService`)—tweak in `getCurrentPrice()` if needed.

Happy tracking & rebalancing! 📈 