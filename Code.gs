// ADD DEBUG FLAG AT TOP
const DEBUG = true;

function ensureSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (headers && sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}

function createSheetsIfNeeded() {
  ensureSheet('TradeList', ['Ticker', 'Price']);
  ensureSheet('Holdings', ['Ticker', 'Quantity', 'AverageCostBasis']);
  ensureSheet('AllocationHistory', ['Date', 'NewTargetAllocation', 'Notes']);
  ensureSheet('TradeHistory', ['Timestamp', 'Action', 'Ticker', 'Quantity', 'Price', 'TotalValue']);
  ensureSheet('Dashboard');
}

function onOpen() {
  createSheetsIfNeeded();
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📈 Portfolio Tools')
    .addItem('▶️ Calculate & Rebalance Portfolio', 'rebalancePortfolio')
    .addToUi();
}

function rebalancePortfolio() {
  logDebug('Rebalance started');
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

  // Phase 1: Gather Data
  const tradeListSheet = ensureSheet('TradeList', ['Ticker', 'Price']);
  const holdingsSheet = ensureSheet('Holdings', ['Ticker', 'Quantity', 'AverageCostBasis']);
  const allocationSheet = ensureSheet('AllocationHistory', ['Date', 'NewTargetAllocation', 'Notes']);
  const tradeHistorySheet = ensureSheet('TradeHistory', ['Timestamp', 'Action', 'Ticker', 'Quantity', 'Price', 'TotalValue']);

  const tradeTickers = tradeListSheet
    .getRange(1, 1, tradeListSheet.getLastRow())
    .getValues()
    .map(r => r[0])
    .filter(Boolean);
  logDebug('Trade tickers', tradeTickers);

  // Ensure price cells for each desired ticker are populated
  tradeTickers.forEach(t => getPriceFromTradeList(t));

  const holdingsData = holdingsSheet
    .getRange(1, 1, holdingsSheet.getLastRow(), 3)
    .getValues()
    .filter(r => r[0]);
  logDebug('Holdings data', holdingsData);

  const latestAllocationRow = allocationSheet.getLastRow();
  const targetAllocation = allocationSheet.getRange(latestAllocationRow, 2).getValue();
  logDebug('Target allocation', targetAllocation);

  if (!tradeTickers.length || !targetAllocation) {
    ui.alert('TradeList or Target Allocation is empty.');
    logDebug('Missing data – aborting');
    return;
  }

  // Phase 2: Calculate Trades
  const allTickers = Array.from(
    new Set([...tradeTickers, ...holdingsData.map(r => r[0])])
  );
  const prices = {};
  allTickers.forEach(t => (prices[t] = getCurrentPrice(t)));
  logDebug('Prices map', prices);

  const targetValuePerTicker = targetAllocation / tradeTickers.length;

  const holdingsMap = {};
  holdingsData.forEach(([ticker, qty, cost]) => {
    holdingsMap[ticker] = { qty, cost };
  });

  const trades = [];

  // Sell tickers not in TradeList
  Object.keys(holdingsMap).forEach(ticker => {
    if (!tradeTickers.includes(ticker)) {
      const qty = holdingsMap[ticker].qty;
      if (qty > 0) trades.push({ action: 'SELL', ticker, quantity: qty, price: prices[ticker] });
    }
  });

  // Adjust positions for desired tickers
  tradeTickers.forEach(ticker => {
    const currentQty = holdingsMap[ticker] ? holdingsMap[ticker].qty : 0;
    const price = prices[ticker];
    const desiredQty = Math.floor(targetValuePerTicker / price);
    const diff = desiredQty - currentQty;
    if (diff > 0) trades.push({ action: 'BUY', ticker, quantity: diff, price });
    else if (diff < 0) trades.push({ action: 'SELL', ticker, quantity: -diff, price });
  });
  logDebug('Trades array', trades);

  if (!trades.length) {
    ui.alert('No trades required. Portfolio already balanced.');
    logDebug('No trades required');
    logDebug('Rebalance complete');
    return;
  }

  // Phase 3: User Confirmation
  let message = 'Proposed Trades:\n\n';
  trades.forEach(t => {
    message += `${t.action} ${t.quantity} ${t.ticker} @ $${t.price.toFixed(2)}\n`;
  });
  logDebug('Proposed trades', trades);
  const response = ui.alert('Review Trades', message, ui.ButtonSet.OK_CANCEL);
  if (response !== ui.Button.OK) {
    ui.alert('Rebalance cancelled.');
    logDebug('User cancelled');
    logDebug('Rebalance complete');
    return;
  }

  // Phase 4: Execute & Record
  const timestamp = new Date();
  trades.forEach(t => {
    tradeHistorySheet.appendRow([
      timestamp,
      t.action,
      t.ticker,
      t.quantity,
      t.price,
      t.quantity * t.price,
    ]);

    const existing = holdingsMap[t.ticker];
    if (t.action === 'BUY') {
      if (existing) {
        const newQty = existing.qty + t.quantity;
        const newCost =
          (existing.qty * existing.cost + t.quantity * t.price) / newQty;
        const rowIdx = holdingsData.findIndex(r => r[0] === t.ticker) + 1;
        holdingsSheet.getRange(rowIdx, 2).setValue(newQty);
        holdingsSheet.getRange(rowIdx, 3).setValue(newCost);
        holdingsMap[t.ticker] = { qty: newQty, cost: newCost };
      } else {
        holdingsSheet.appendRow([t.ticker, t.quantity, t.price]);
        holdingsMap[t.ticker] = { qty: t.quantity, cost: t.price };
      }
    } else {
      // SELL
      if (!existing) return;
      const newQty = existing.qty - t.quantity;
      const rowIdx = holdingsData.findIndex(r => r[0] === t.ticker) + 1;
      if (newQty <= 0) {
        holdingsSheet.deleteRow(rowIdx);
        delete holdingsMap[t.ticker];
      } else {
        holdingsSheet.getRange(rowIdx, 2).setValue(newQty);
        holdingsMap[t.ticker] = { ...existing, qty: newQty };
      }
    }
  });
  logDebug('Trades executed');

  // Phase 5: Update Dashboard
  updateDashboard();
  ui.alert('Rebalance complete.');
  logDebug('Rebalance complete');
  } catch (e) {
    logDebug('ERROR', e);
    throw e;
  }
}

function updateDashboard() {
  logDebug('updateDashboard start');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ensureSheet('Dashboard');
  const holdingsSheet = ensureSheet('Holdings', ['Ticker', 'Quantity', 'AverageCostBasis']);
  const allocationSheet = ensureSheet('AllocationHistory', ['Date', 'NewTargetAllocation', 'Notes']);
  if (!dashboardSheet || !holdingsSheet || !allocationSheet) return;

  const holdingsData = holdingsSheet
    .getRange(1, 1, holdingsSheet.getLastRow(), 3)
    .getValues()
    .filter(r => r[0]);
  const targetAllocation = allocationSheet.getRange(allocationSheet.getLastRow(), 2).getValue();

  if (!holdingsData.length) {
    dashboardSheet.clear();
    return;
  }

  const prices = {};
  holdingsData.forEach(([ticker]) => (prices[ticker] = getCurrentPrice(ticker)));

  let portfolioValue = 0;
  let costBasisTotal = 0;
  const tableRows = holdingsData.map(([ticker, qty, avgCost]) => {
    const price = prices[ticker];
    const marketValue = qty * price;
    const costBasis = qty * avgCost;
    const pnl = marketValue - costBasis;
    const pnlPct = costBasis ? pnl / costBasis : 0;

    portfolioValue += marketValue;
    costBasisTotal += costBasis;

    return [ticker, qty, avgCost, price, marketValue, pnlPct, pnl];
  });

  const cashPosition = targetAllocation - costBasisTotal;
  const totalPnl = portfolioValue - costBasisTotal;

  // Key metrics
  dashboardSheet.getRange('B1').setValue(targetAllocation);
  dashboardSheet.getRange('B2').setValue(portfolioValue);
  dashboardSheet.getRange('B3').setValue(cashPosition);
  dashboardSheet.getRange('B4').setValue(totalPnl);

  // Table headers & data
  dashboardSheet.getRange('A10:Z1000').clear();
  const headers = [
    'Ticker',
    'Quantity',
    'Avg. Trade Price',
    'Current Price',
    'Market Value',
    'P&L %',
    'P&L Open ($)',
  ];
  dashboardSheet.getRange(10, 1, 1, headers.length).setValues([headers]);
  dashboardSheet
    .getRange(11, 1, tableRows.length, headers.length)
    .setValues(tableRows);
  logDebug('updateDashboard done');
}

// Helper: ensure price in TradeList and return it
function getPriceFromTradeList(ticker) {
  const sheet = ensureSheet('TradeList', ['Ticker', 'Price']);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === ticker) {
      const rowIdx = i + 2; // header offset
      let price = data[i][1];
      if (!price) {
        const priceCell = sheet.getRange(rowIdx, 2);
        priceCell.setFormula(`=GOOGLEFINANCE(A${rowIdx},"price")`);
        SpreadsheetApp.flush();
        Utilities.sleep(1500);
        price = priceCell.getValue();
      }
      return parseFloat(price);
    }
  }
  return null;
}

function getCurrentPrice(ticker) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(ticker);
  if (cached) return parseFloat(cached);

  // First attempt to use TradeList price column
  let price = getPriceFromTradeList(ticker);
  if (price !== null && !isNaN(price)) {
    cache.put(ticker, price, 300);
    return price;
  }

  // Fallback: temporarily insert row to fetch price
  const tradeListSheet = ensureSheet('TradeList', ['Ticker', 'Price']);
  const tempRow = tradeListSheet.getLastRow() + 1;
  tradeListSheet.getRange(tempRow, 1).setValue(ticker);
  const priceCell = tradeListSheet.getRange(tempRow, 2);
  priceCell.setFormula(`=GOOGLEFINANCE(A${tempRow},"price")`);
  SpreadsheetApp.flush();
  Utilities.sleep(1500);
  price = parseFloat(priceCell.getValue());
  tradeListSheet.deleteRow(tempRow);

  cache.put(ticker, price, 300); // 5-min cache
  return price;
}

function logDebug(msg, payload) {
  if (!DEBUG) return;
  if (payload !== undefined) {
    Logger.log(msg + ': ' + JSON.stringify(payload));
  } else {
    Logger.log(msg);
  }
}

// ADD RUN WRAPPER AT END
function runRebalance() {
  return rebalancePortfolio();
}
