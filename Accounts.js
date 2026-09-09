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