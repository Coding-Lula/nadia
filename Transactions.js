function addTransaction(txPayload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const txSheet = ss.getSheetByName('Transactions');
  const accSheet = ss.getSheetByName('Accounts');
  
  const txId = generateId('TX');
  const dateStr = txPayload.date || Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
  const date = new Date(dateStr + "T00:00:00");
  const { accountId, type, category, amount, description, targetAccountId, isDebt, debtId } = txPayload;
  const numAmount = Number(amount);
  
  const todayStr = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
  const isFuture = dateStr > todayStr;

  let finalCategory = category;
  let finalDescription = description || '';

  if (isDebt && debtId) {
    finalCategory = 'Amortizar Divida';
    const debtInfo = recordDebtPayment(debtId, numAmount);
    if (debtInfo && debtInfo.person) {
      finalDescription = finalDescription ? `${finalDescription} (Dívida: ${debtInfo.person})` : `Amortização de dívida: ${debtInfo.person}`;
    }
  }

  // Columns: [0: ID, 1: Date, 2: Account_ID, 3: Type, 4: Category, 5: Amount, 6: Description, 7: TargetAccount_ID, 8: Debt_ID]
  txSheet.appendRow([txId, date, accountId, type, finalCategory, numAmount, finalDescription, targetAccountId || '', debtId || '']);
  
  // Sync Accounts Sheet balances ONLY if transaction date <= today
  if (!isFuture) {
    updateAccountBalance(accSheet, accountId, type, numAmount, true);
    if (type === 'Transfer' && targetAccountId) {
      updateAccountBalance(accSheet, targetAccountId, 'Income', numAmount, false);
    }
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

// Recalculates effective account balance taking into account transactions up to today
function calculateAccountBalanceUpToToday(accountId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const accSheet = ss.getSheetByName('Accounts');
  const txSheet = ss.getSheetByName('Transactions');

  if (!accSheet) return 0;

  // Find initial balance
  let initialBalance = 0;
  const accData = accSheet.getDataRange().getValues();
  for (let i = 1; i < accData.length; i++) {
    if (accData[i][0].toString() === accountId.toString()) {
      initialBalance = Number(accData[i][3]) || 0; // Column 3: initial balance
      break;
    }
  }

  if (!txSheet) return initialBalance;

  const todayStr = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
  const txData = txSheet.getDataRange().getValues();

  let balance = initialBalance;
  for (let i = 1; i < txData.length; i++) {
    const row = txData[i];
    if (!row[0]) continue;

    const rowDate = row[1] ? Utilities.formatDate(new Date(row[1]), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd") : '';
    if (rowDate > todayStr) continue; // Skip future transactions

    const srcAcc = row[2] ? row[2].toString() : '';
    const type = row[3];
    const amount = Number(row[5]) || 0;
    const targetAcc = row[7] ? row[7].toString() : '';

    if (srcAcc === accountId.toString()) {
      if (type === 'Income') balance += amount;
      else if (type === 'Expense' || type === 'Transfer') balance -= amount;
    } else if (targetAcc === accountId.toString() && type === 'Transfer') {
      balance += amount;
    }
  }

  return balance;
}

// Fixes the 0,00 MT issue on Painel Geral summary boxes for current month
function getCurrentMonthSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Transactions');
  if (!sheet) return { monthlyIncome: 0, monthlyExpense: 0 };
  const data = sheet.getDataRange().getValues();
  
  const now = new Date();
  const currentMonth = now.getMonth();
  const currentYear = now.getFullYear();
  const todayStr = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
  
  let monthlyIncome = 0;
  let monthlyExpense = 0;
  
  for (let i = 1; i < data.length; i++) {
    const [txId, txDate, accId, type, category, amount] = data[i];
    if (!txDate) continue;
    const rowDateObj = new Date(txDate);
    const rowDateStr = Utilities.formatDate(rowDateObj, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
    
    // Only count transactions up to today for current month
    if (rowDateObj.getMonth() === currentMonth && rowDateObj.getFullYear() === currentYear && rowDateStr <= todayStr) {
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
    const rawAmt = parseFloat(row[5]) || 0;
    const type = row[3];
    const displayAmount = (type === 'Expense') ? -Math.abs(rawAmt) : Math.abs(rawAmt);

    return {
      id: row[0],
      date: row[1] ? Utilities.formatDate(new Date(row[1]), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd") : '',
      accountId: accId,
      accountName: accMap[accId] || accId || 'N/A',
      type: type,
      category: row[4],
      amount: displayAmount,
      description: row[6] || '',
      targetAccountId: row[7] || '',
      debtId: row[8] || ''
    };
  }).reverse();
}

/**
 * Delete transaction by ID and adjust account & debt balances accordingly
 */
function deleteTransaction(txId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Transactions');
  const accSheet = ss.getSheetByName('Accounts');
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  const todayStr = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");

  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === txId.toString()) {
      const row = data[i];
      const txDateStr = row[1] ? Utilities.formatDate(new Date(row[1]), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd") : '';
      const isFuture = txDateStr > todayStr;
      const accountId = row[2];
      const type = row[3];
      const numAmount = Number(row[5]) || 0;
      const targetAccountId = row[7];
      const debtId = row[8];

      // Revert account balances if transaction was past/present (not future)
      if (!isFuture) {
        if (type === 'Expense') {
          // Refund expense back to source account
          updateAccountBalance(accSheet, accountId, 'Income', numAmount, false);
        } else if (type === 'Income') {
          // Deduct income from source account
          updateAccountBalance(accSheet, accountId, 'Expense', numAmount, false);
        } else if (type === 'Transfer') {
          // Reverse transfer: return to source, deduct from target
          updateAccountBalance(accSheet, accountId, 'Income', numAmount, false);
          if (targetAccountId) {
            updateAccountBalance(accSheet, targetAccountId, 'Expense', numAmount, false);
          }
        }
      }

      // Revert debt payment if debtId exists
      if (debtId) {
        revertDebtPayment(debtId, numAmount);
      }

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
  const todayStr = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");

  const categoryTotals = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[1]) continue;
    const txDate = new Date(row[1]);
    const rowDateStr = Utilities.formatDate(txDate, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
    const type = row[3];
    const category = row[4] || 'Outros';
    const amount = parseFloat(row[5]) || 0;

    if (
      txDate.getMonth() === currentMonth &&
      txDate.getFullYear() === currentYear &&
      rowDateStr <= todayStr &&
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