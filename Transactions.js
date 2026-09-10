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

// Fixes the 0,00 MT issue on Painel Geral summary boxes for current month
function getCurrentMonthSummary() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
  const data = sheet.getDataRange().getValues();
  
  const now = new Date();
  const currentMonth = now.getMonth();
  const currentYear = now.getFullYear();
  
  let monthlyIncome = 0;
  let monthlyExpense = 0;
  
  for (let i = 1; i < data.length; i++) {
    const [txId, txDate, accId, type, category, amount] = data[i];
    const rowDate = new Date(txDate);
    
    if (rowDate.getMonth() === currentMonth && rowDate.getFullYear() === currentYear) {
      const val = Number(amount) || 0;
      if (type === 'Income') monthlyIncome += val;
      if (type === 'Expense') monthlyExpense += val;
    }
  }
  
  return { monthlyIncome, monthlyExpense };
}
/**
 * Fetches recent transactions across all accounts for the Global Transactions tab
 */
function getRecentTransactions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const txSheet = ss.getSheetByName('Transactions');
  const accSheet = ss.getSheetByName('Accounts');

  if (!txSheet) return [];

  // 1. Build Account ID -> Account Name map
  const accMap = {};
  if (accSheet) {
    const accRows = accSheet.getDataRange().getValues().slice(1);
    accRows.forEach(row => {
      accMap[row[0]] = row[1]; // row[0] = Account_ID, row[1] = Account_Name
    });
  }

  const rows = txSheet.getDataRange().getValues().slice(1);

  // Return mapped transaction objects including accountName
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
      // Adjust column indices to match your sheet layout:
      // Row structure assumed: [ID, Date, Account, Type, Category, Amount, Description]
      sheet.getRange(i + 1, 5).setValue(payload.category); // Category Column
      sheet.getRange(i + 1, 6).setValue(parseFloat(payload.amount)); // Amount Column
      sheet.getRange(i + 1, 7).setValue(payload.description); // Description Column
      break;
    }
  }
}
/**
 * Calculates current month's expenses aggregated by category for the chart
 */
function getCategoryExpenseBreakdown() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Transactions');
  if (!sheet) return { labels: [], values: [] };

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { labels: [], values: [] };

  const now = new Date();
  const currentMonth = now.getMonth();
  const currentYear = now.getFullYear();

  const categoryTotals = {};

  // Assuming columns: [0: ID, 1: Date, 2: Account, 3: Type, 4: Category, 5: Amount, 6: Description]
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const txDate = new Date(row[1]);
    const type = row[3];
    const category = row[4] || 'Outros';
    const amount = parseFloat(row[5]) || 0;

    // Filter for Expenses in the current month
    if (
      txDate.getMonth() === currentMonth &&
      txDate.getFullYear() === currentYear &&
      (type === 'Expense' || amount < 0)
    ) {
      const positiveAmt = Math.abs(amount);
      categoryTotals[category] = (categoryTotals[category] || 0) + positiveAmt;
    }
  }

  return {
    labels: Object.keys(categoryTotals),
    values: Object.values(categoryTotals)
  };
}