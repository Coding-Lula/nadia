function addTransaction(txPayload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const txSheet = ss.getSheetByName('Transactions');
  const accSheet = ss.getSheetByName('Accounts');
  
  const txId = generateId('TX');
  const date = txPayload.date ? new Date(txPayload.date) : new Date();
  const { accountId, type, category, amount, description, targetAccountId } = txPayload;
  const numAmount = Number(amount);
  
  txSheet.appendRow([txId, date, accountId, type, category, numAmount, description, targetAccountId || '']);
  
  // Sync Accounts Sheet balances
  updateAccountBalance(accSheet, accountId, type, numAmount, true);
  if (type === 'Transfer' && targetAccountId) {
    updateAccountBalance(accSheet, targetAccountId, 'Income', numAmount, false);
  }
  
  return { status: 'SUCCESS' };
}

function updateAccountBalance(accSheet, accountId, type, amount, isSource) {
  const data = accSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == accountId) {
      let currentBalance = Number(data[i][4]) || 0;
      if (type === 'Income') currentBalance += amount;
      else if (type === 'Expense') currentBalance -= amount;
      else if (type === 'Transfer' && isSource) currentBalance -= amount;
      
      accSheet.getRange(i + 1, 5).setValue(currentBalance);
      break;
    }
  }
}

// Summary for current month and comparison with previous month
function getCurrentMonthSummary() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
  if (!sheet) return { monthlyIncome: 0, monthlyExpense: 0, netChangePercent: 0, vsLastMonth: 0 };

  const data = sheet.getDataRange().getValues();
  
  const now = new Date();
  const currentMonth = now.getMonth();
  const currentYear = now.getFullYear();
  
  const prevMonthDate = new Date(currentYear, currentMonth - 1, 1);
  const prevMonth = prevMonthDate.getMonth();
  const prevYear = prevMonthDate.getFullYear();

  let monthlyIncome = 0;
  let monthlyExpense = 0;
  let prevMonthExpense = 0;
  let prevMonthIncome = 0;
  
  for (let i = 1; i < data.length; i++) {
    const [txId, txDate, accId, type, category, amount] = data[i];
    if (!txDate) continue;
    const rowDate = new Date(txDate);
    const val = Number(amount) || 0;
    
    if (rowDate.getMonth() === currentMonth && rowDate.getFullYear() === currentYear) {
      if (type === 'Income') monthlyIncome += val;
      if (type === 'Expense') monthlyExpense += val;
    } else if (rowDate.getMonth() === prevMonth && rowDate.getFullYear() === prevYear) {
      if (type === 'Income') prevMonthIncome += val;
      if (type === 'Expense') prevMonthExpense += val;
    }
  }
  
  let vsLastMonth = 0;
  if (prevMonthExpense > 0) {
    vsLastMonth = ((monthlyExpense - prevMonthExpense) / prevMonthExpense) * 100;
  } else if (monthlyExpense > 0) {
    vsLastMonth = 100;
  }

  return {
    monthlyIncome,
    monthlyExpense,
    prevMonthExpense,
    prevMonthIncome,
    vsLastMonth: Number(vsLastMonth.toFixed(1))
  };
}

/**
 * Fetches recent transactions across all accounts
 */
function getRecentTransactions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const txSheet = ss.getSheetByName('Transactions');
  const accSheet = ss.getSheetByName('Accounts');

  if (!txSheet) return [];

  // Build Account ID -> Account Name map
  const accMap = {};
  if (accSheet) {
    const accRows = accSheet.getDataRange().getValues().slice(1);
    accRows.forEach(row => {
      accMap[row[0]] = row[1];
    });
  }

  const rows = txSheet.getDataRange().getValues().slice(1);

  return rows.map(row => {
    const accId = row[2];
    return {
      id: row[0],
      date: row[1] ? Utilities.formatDate(new Date(row[1]), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd") : '',
      accountId: accId,
      accountName: accMap[accId] || accId || 'N/A',
      type: row[3],
      category: row[4],
      amount: parseFloat(row[5]) || 0,
      description: row[6] || ''
    };
  }).reverse();
}

/**
 * Delete transaction by ID
 */
function deleteTransaction(txId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Transactions');
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === txId.toString()) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
}

/**
 * Update transaction details
 */
function updateTransaction(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Transactions');
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === payload.id.toString()) {
      sheet.getRange(i + 1, 5).setValue(payload.category);
      sheet.getRange(i + 1, 6).setValue(parseFloat(payload.amount));
      sheet.getRange(i + 1, 7).setValue(payload.description);
      break;
    }
  }
}

/**
 * Calculates current month's breakdown aggregated by category for chart and list.
 * @param {string} filterType 'Expense' or 'Income'
 */
function getCategoryBreakdown(filterType) {
  const targetType = filterType === 'Income' ? 'Income' : 'Expense';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Transactions');
  if (!sheet) return { labels: [], values: [], total: 0, items: [] };

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { labels: [], values: [], total: 0, items: [] };

  const now = new Date();
  const currentMonth = now.getMonth();
  const currentYear = now.getFullYear();

  const categoryTotals = {};
  let totalSum = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[1]) continue;
    const txDate = new Date(row[1]);
    const type = row[3];
    const category = row[4] || 'Outros';
    const amount = parseFloat(row[5]) || 0;

    if (
      txDate.getMonth() === currentMonth &&
      txDate.getFullYear() === currentYear &&
      type === targetType
    ) {
      const positiveAmt = Math.abs(amount);
      categoryTotals[category] = (categoryTotals[category] || 0) + positiveAmt;
      totalSum += positiveAmt;
    }
  }

  // Convert to sorted array of items descending by amount
  const sortedCategories = Object.keys(categoryTotals).map(cat => ({
    category: cat,
    amount: categoryTotals[cat],
    percentage: totalSum > 0 ? (categoryTotals[cat] / totalSum) * 100 : 0
  })).sort((a, b) => b.amount - a.amount);

  return {
    labels: sortedCategories.map(c => c.category),
    values: sortedCategories.map(c => c.amount),
    total: totalSum,
    items: sortedCategories
  };
}

// Backward-compatible alias for existing callers
function getCategoryExpenseBreakdown() {
  return getCategoryBreakdown('Expense');
}
