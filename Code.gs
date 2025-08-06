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

  const tradeLastRow = tradeListSheet.getLastRow();
  const tradeTickers = tradeLastRow > 1
    ? tradeListSheet
        .getRange(2, 1, tradeLastRow - 1)
        .getValues()
        .map(r => r[0])
        .filter(Boolean)
    : [];
  logDebug('Trade tickers', tradeTickers);

  // Ensure price cells for each desired ticker are populated
  tradeTickers.forEach(t => getPriceFromTradeList(t));

  const holdingsLastRow = holdingsSheet.getLastRow();
  const holdingsData = holdingsLastRow > 1
    ? holdingsSheet
        .getRange(2, 1, holdingsLastRow - 1, 3)
        .getValues()
        .filter(r => r[0])
    : [];
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
  const forcedOneShare = [];
  const noPrice = [];

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
    if (price == null || isNaN(price)) {
      noPrice.push(ticker);
      return;
    }
    const desiredQty = Math.floor(targetValuePerTicker / price);
    let finalDesiredQty = desiredQty;
    if (finalDesiredQty === 0) {
      finalDesiredQty = 1;
      forcedOneShare.push(ticker);
    }
    const diff = finalDesiredQty - currentQty;
    if (diff > 0) trades.push({ action: 'BUY', ticker, quantity: diff, price });
    else if (diff < 0) trades.push({ action: 'SELL', ticker, quantity: -diff, price });
  });
  logDebug('Trades array', trades);
  logDebug('Forced 1-share tickers', forcedOneShare);
  logDebug('Tickers with no price', noPrice);

  if (!trades.length) {
    ui.alert('No trades required. Portfolio already balanced.');
    logDebug('No trades required');
    logDebug('Rebalance complete');
    return;
  }

  // Prepare summary list for executed trades (filled later)
  let executedSection = '';
  trades.forEach(t => {
    executedSection += `${t.action} ${t.quantity} ${t.ticker} @ $${t.price.toFixed(2)}\n`;
  });
  logDebug('Trades to execute', trades);

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
        const rowIdx = holdingsData.findIndex(r => r[0] === t.ticker) + 2; // adjust for header
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
      const rowIdx = holdingsData.findIndex(r => r[0] === t.ticker) + 2; // adjust for header
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

  // Assemble final summary message
  let summaryMessage = 'Executed Trades:\n\n' + (executedSection || 'None') + '\n';
  if (forcedOneShare.length) {
    summaryMessage += '\nTickers defaulted to 1 share due to high price:\n' + forcedOneShare.join(', ') + '\n';
  }
  if (noPrice.length) {
    summaryMessage += '\nTickers skipped due to missing price:\n' + noPrice.join(', ') + '\n';
  }

  ui.alert('Rebalance Summary', summaryMessage, ui.ButtonSet.OK);
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

  // Fetch holdings (skip header)
  const hLast = holdingsSheet.getLastRow();
  const holdingsData = hLast > 1
    ? holdingsSheet.getRange(2, 1, hLast - 1, 3).getValues().filter(r => r[0])
    : [];

  // Reset dashboard
  dashboardSheet.clear();

  // Write table headers
  const headers = [
    'Ticker',
    'Quantity',
    'Avg. Trade Price',
    'Current Price',
    'Market Value',
    'P&L %',
    'P&L Open ($)',
  ];
  dashboardSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');

  // Basic holdings rows (Ticker, Qty, AvgCost)
  if (holdingsData.length) {
    dashboardSheet.getRange(2, 1, holdingsData.length, 3).setValues(holdingsData);
  }

  // Insert ARRAYFORMULAS for dynamic columns
  if (holdingsData.length) {
    dashboardSheet.getRange('D2').setFormula(
      '=ARRAYFORMULA(IF(A2:A="",,IFERROR(VLOOKUP(A2:A,TradeList!A:B,2,FALSE),GOOGLEFINANCE(A2:A,"price"))))'
    );
    dashboardSheet.getRange('E2').setFormula(
      '=ARRAYFORMULA(IF(A2:A="",,B2:B*D2:D))'
    );
    dashboardSheet.getRange('G2').setFormula(
      '=ARRAYFORMULA(IF(A2:A="",,E2:E-B2:B*C2:C))'
    );
    dashboardSheet.getRange('F2').setFormula(
      '=ARRAYFORMULA(IF(A2:A="",,IF(C2:C=0,0,G2:G/(B2:B*C2:C))))'
    );
  }

  // Key metrics labels and formulas on the right
  const metricLabels = [
    ['Target Allocation'],
    ['Portfolio Value'],
    ['Cash Position'],
    ['Unrealized P&L'],
    ['Total P&L'],
    ['Starting Capital'],
    ['YTD P&L'],
    ['YTD P&L %'],
    ['Allocation / Stock'],
  ];
  dashboardSheet.getRange('J1:J9').setValues(metricLabels).setFontWeight('bold');

  dashboardSheet.getRange('K1').setFormula('=INDEX(AllocationHistory!B:B,COUNTA(AllocationHistory!B:B))');
  dashboardSheet.getRange('K2').setFormula('=SUM(E2:E)');
  dashboardSheet.getRange('K3').setFormula('=K1-SUM(B2:B*C2:C)');
  dashboardSheet.getRange('K4').setFormula('=K2-SUM(B2:B*C2:C)');
  dashboardSheet.getRange('K5').setFormula('=K2-INDEX(AllocationHistory!B:B, MATCH(TRUE, ISNUMBER(AllocationHistory!B:B),0))');
  dashboardSheet.getRange('K6').setFormula('=IFERROR(INDEX(FILTER(AllocationHistory!B:B, YEAR(AllocationHistory!A:A)=YEAR(TODAY())),1),0)');
  dashboardSheet.getRange('K7').setFormula('=K2-K6');
  dashboardSheet.getRange('K8').setFormula('=IF(K6=0,0,K7/K6)');
  dashboardSheet.getRange('K9').setFormula('=K1/(COUNTA(TradeList!A:A)-1)');

  // Formatting
  const lastRow = holdingsData.length + 1; // including header
  dashboardSheet.getRange(2, 3, Math.max(holdingsData.length,1), 1).setNumberFormat('$#,##0.00'); // Avg Trade Price
  dashboardSheet.getRange(2, 4, Math.max(holdingsData.length,1), 1).setNumberFormat('$#,##0.00'); // Current Price
  dashboardSheet.getRange(2, 5, Math.max(holdingsData.length,1), 1).setNumberFormat('$#,##0.00'); // Market Value
  dashboardSheet.getRange(2, 7, Math.max(holdingsData.length,1), 1).setNumberFormat('$#,##0.00'); // P&L Open ($)
  dashboardSheet.getRange(2, 6, Math.max(holdingsData.length,1), 1).setNumberFormat('0.00%'); // P&L %
  dashboardSheet.getRange('K1:K7').setNumberFormat('$#,##0.00');
  dashboardSheet.getRange('K9').setNumberFormat('$#,##0.00');
  dashboardSheet.getRange('K8').setNumberFormat('0.00%');

  // Auto resize columns for better visibility
  dashboardSheet.autoResizeColumns(1, 7); // A:G
  dashboardSheet.autoResizeColumns(10, 2); // J:K

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
