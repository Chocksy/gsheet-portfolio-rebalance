[MODE: PLAN]

This is a fantastic clarification. The manual, signal-driven approach changes the core logic significantly, making the spreadsheet's primary role to track state and calculate trades based on your deliberate decisions, rather than automatically rebalancing. The "Target Allocation" becomes a key, manually-adjusted variable.

This is a much more robust and controllable system. Here is the detailed plan to build it.

### Spreadsheet Structure Plan

1.  **Create a New Google Spreadsheet** and name it "Portfolio Tracker".
2.  **Create Sheet: `Dashboard`**
    *   This will be the main view. It will be read-only for the user, as all its data is calculated.
    *   **Key Values (Top Section):**
        *   `B1`: **Target Allocation:** (e.g., $20,000) - This value will be pulled from the `AllocationHistory` sheet.
        *   `B2`: **Current Holdings Value:** (Sum of all positions at current market price).
        *   `B3`: **Cash Position:** (`Target Allocation` - `Cost of Current Holdings`).
        *   `B4`: **Total Unrealized P&L:** (Sum of P&L for all open positions).
    *   **Holdings Table (Main Section):**
        *   Columns: `Ticker`, `Quantity`, `Avg. Trade Price`, `Current Price`, `Market Value`, `P&L %`, `P&L Open ($)`. This table will be populated automatically by the script.
3.  **Create Sheet: `TradeList`** (User Input)
    *   This is the *only* sheet (besides `AllocationHistory`) the user should manually edit.
    *   It represents the desired state of the portfolio.
    *   **Column A:** `Ticker`. The user will add or remove tickers from this single column to signal their intent for the next rebalance.
4.  **Create Sheet: `Holdings`** (Source of Truth)
    *   This sheet represents the *actual* current state of the portfolio. It is managed entirely by the script.
    *   **Column A:** `Ticker`
    *   **Column B:** `Quantity` (Whole number of shares)
    *   **Column C:** `AverageCostBasis` (The weighted average price paid for the shares held).
5.  **Create Sheet: `TradeHistory`** (Append-Only Log)
    *   The script will log every transaction here.
    *   **Columns:** `Timestamp`, `Action` (BUY/SELL), `Ticker`, `Quantity`, `Price`, `TotalValue`.
6.  **Create Sheet: `AllocationHistory`** (User Input)
    *   The user logs any changes to their total desired investment here.
    *   **Column A:** `Date`
    *   **Column B:** `NewTargetAllocation`
    *   **Column C:** `Notes` (e.g., "Initial funding", "Added Q1 profits").
    *   The script will always use the *last entry* in this sheet as the official "Target Allocation".

### Google Apps Script (`Code.gs`) Plan

1.  **`onOpen()` Function:**
    *   Runs when the spreadsheet is opened.
    *   Creates a custom menu named "📈 Portfolio Tools" in the UI.
    *   This menu will have one option: "▶️ Calculate & Rebalance Portfolio".
2.  **`rebalancePortfolio()` Function (The Core Logic):**
    *   This is the main function, triggered from the menu. It will orchestrate the entire process.
    *   **Phase 1: Gather Data**
        1.  Read the desired tickers from the `TradeList` sheet.
        2.  Read the current holdings (ticker, quantity, cost basis) from the `Holdings` sheet.
        3.  Read the latest `Target Allocation` amount from the `AllocationHistory` sheet.
    *   **Phase 2: Calculate Trades**
        1.  Determine the target value for each stock (`Target Allocation` / number of tickers in `TradeList`).
        2.  Fetch current market prices for all relevant tickers (both held and desired) using the `GOOGLEFINANCE` function.
        3.  Iterate through the `TradeList` and `Holdings` to generate a list of proposed trades:
            *   **Identify Stocks to SELL entirely:** Tickers in `Holdings` but not in `TradeList`.
            *   **Identify Stocks to BUY/SELL to adjust:** For each ticker in `TradeList`, calculate its target share count (`target value` / `current price`, rounded down). Compare this to the current share count and generate a BUY or SELL trade for the difference.
            *   **Identify new Stocks to BUY:** Tickers in `TradeList` but not in `Holdings`.
    *   **Phase 3: User Confirmation**
        1.  Format the list of proposed trades into a clear, human-readable message.
        2.  Display this message in a pop-up confirmation dialog (`SpreadsheetApp.getUi().alert()`) that shows every planned BUY and SELL action.
        3.  The dialog will have "Execute Trades" and "Cancel" buttons. The script will halt if the user cancels.
    *   **Phase 4: Execute & Record (If Confirmed)**
        1.  For each confirmed trade, append a new row to the `TradeHistory` sheet.
        2.  Update the `Holdings` sheet: modify quantities, add new stocks with their cost basis, and remove stocks that were sold completely. Recalculate the `AverageCostBasis` for any positions that were added to.
    *   **Phase 5: Update Dashboard**
        1.  Call a final function, `updateDashboard()`, to refresh all calculated values and tables in the `Dashboard` sheet.
        2.  Show a "Rebalance Complete" confirmation message.
3.  **`updateDashboard()` Helper Function:**
    *   This function reads the updated `Holdings` sheet, fetches live prices, and populates the `Dashboard` with all the current portfolio metrics (Market Value, P&L, etc.).

I will now convert this plan into a step-by-step todo list in your Cursor IDE panel.