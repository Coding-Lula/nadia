function calculateAccountBalanceUpToToday(accountId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const accSheet = ss.getSheetByName('Accounts');
  if (!accSheet) return 0;

  const accData = accSheet.getDataRange().getValues();
  let balance = 0;

  for (let i = 1; i < accData.length; i++) {
    if (accData[i][0] && String(accData[i][0]) === String(accountId)) {
      balance = Number(accData[i][3]) || 0;
      break;
    }
  }

  const txSheet = ss.getSheetByName('Transactions');
  if (!txSheet) return balance;

  const txData = txSheet.getDataRange().getValues();
  const now = new Date();
  now.setHours(23, 59, 59, 999);

  for (let i = 1; i < txData.length; i++) {
    const [txId, txDate, sourceAccId, type, category, rawAmount, description, targetAccId] = txData[i];
    if (!txDate) continue;

    const rowDate = new Date(txDate);
    if (rowDate > now) continue;

    const amount = Math.abs(Number(rawAmount) || 0);

    const isSource = sourceAccId && String(sourceAccId) === String(accountId);
    const isTarget = targetAccId && String(targetAccId) === String(accountId);

    if (isSource) {
      if (type === 'Income') balance += amount;
      else if (type === 'Expense') balance -= amount;
      else if (type === 'Transfer') balance -= amount;
    } else if (isTarget) {
      if (type === 'Transfer') balance += amount;
    }
  }

  return balance;
}

function getAccountsOverview() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Accounts');
  if (!sheet) return { totalBalance: 0, accounts: [] };
  const data = sheet.getDataRange().getValues();
  
  const accounts = [];
  let grandTotal = 0;
  
  for (let i = 1; i < data.length; i++) {
    const [id, name, type, initialBal, currentBal] = data[i];
    if (!id) continue;
    // Calculate balance dynamically taking into account transactions up to today
    const balance = calculateAccountBalanceUpToToday(id);
    grandTotal += balance;
    accounts.push({ id, name, type, initialBalance: Number(initialBal), balance });
  }
  
  return { totalBalance: grandTotal, accounts };
}

function addAccount(name, type, initialBalance) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Accounts');
  const id = generateId('ACC');
  const initBal = Number(initialBalance) || 0;
  sheet.appendRow([id, name, type, initBal, initBal]);
  return { status: 'SUCCESS', id };
}

function updateAccount(id, name, type) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Accounts');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.getRange(i + 1, 2).setValue(name);
      sheet.getRange(i + 1, 3).setValue(type);
      return { status: 'SUCCESS' };
    }
  }
  return { status: 'ERROR', message: 'Account not found' };
}

function deleteAccount(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Accounts');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.deleteRow(i + 1);
      return { status: 'SUCCESS' };
    }
  }
  return { status: 'ERROR', message: 'Account not found' };
}