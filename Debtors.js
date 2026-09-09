function getDebts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Debts');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const debts = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      debts.push({
        id: data[i][0],
        person: data[i][1],
        type: data[i][2], // 'Lend' (Quem me deve) or 'Borrow' (A quem devo)
        totalAmount: Number(data[i][3]),
        paidAmount: Number(data[i][4]),
        dueDate: data[i][5] ? Utilities.formatDate(new Date(data[i][5]), Session.getScriptTimeZone(), "yyyy-MM-dd") : '-'
      });
    }
  }
  return debts;
}

function addDebt(person, type, amount, dueDate) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Debts');
  const id = generateId('DEBT');
  sheet.appendRow([id, person, type, Number(amount), 0, dueDate]);
  return { status: 'SUCCESS' };
}

function recordDebtPayment(debtId, paymentAmount) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Debts');
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() == debtId.toString()) {
      const currentPaid = Number(data[i][4]) || 0;
      const newPaid = currentPaid + Number(paymentAmount);
      sheet.getRange(i + 1, 5).setValue(newPaid);
      return { status: 'SUCCESS', person: data[i][1], newPaid: newPaid };
    }
  }
  return null;
}

function revertDebtPayment(debtId, paymentAmount) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Debts');
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() == debtId.toString()) {
      const currentPaid = Number(data[i][4]) || 0;
      const newPaid = Math.max(0, currentPaid - Number(paymentAmount));
      sheet.getRange(i + 1, 5).setValue(newPaid);
      return { status: 'SUCCESS', person: data[i][1], newPaid: newPaid };
    }
  }
  return null;
}