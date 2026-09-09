/**
 * Calculates account statement totals handling positive transaction amounts paired with Type ("Income" / "Expense")
 */
function getStatementData(accId, start, end, type, category) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const txSheet = ss.getSheetByName('Transactions');
  const accSheet = ss.getSheetByName('Accounts');

  if (!txSheet || !accSheet) return { openingBalance: 0, totalIncome: 0, totalExpense: 0, closingBalance: 0, transactions: [] };

  // 1. Map Account_ID to Account_Name
  const accMap = {};
  const accRows = accSheet.getDataRange().getValues().slice(1);
  accRows.forEach(row => {
    accMap[row[0]] = row[1];
  });

  const startDate = new Date(start);
  const endDate = new Date(end);
  endDate.setHours(23, 59, 59, 999);

  const txData = txSheet.getDataRange().getValues();
  const rows = txData.slice(1);

  let openingBalance = 0;
  let totalIncome = 0;
  let totalExpense = 0;
  const filteredTransactions = [];

  rows.forEach(row => {
    // Columns: [0: Tx_ID, 1: Date, 2: Account_ID, 3: Type, 4: Category, 5: Amount, 6: Description]
    const rowAccId = row[2] ? row[2].toString() : '';
    const txDate = new Date(row[1]);
    const txType = row[3]; // 'Income', 'Expense', or 'Transfer'
    const txCategory = row[4] || '';
    const rawAmount = parseFloat(row[5]) || 0;
    const amount = Math.abs(rawAmount); // Standardize to positive number

    const matchesAccount = (accId === 'ALL' || accId === '' || !accId) ? true : (rowAccId === accId.toString());

    if (matchesAccount) {
      
      // A. Calculate Opening Balance (Transactions prior to start date)
      if (txDate < startDate) {
        if (txType === 'Income') {
          openingBalance += amount;
        } else if (txType === 'Expense') {
          openingBalance -= amount;
        }
      }

      // B. Filter Transactions within Date Range
      if (txDate >= startDate && txDate <= endDate) {
        const matchesType = (type === 'ALL' || type === '' || !type) ? true : (txType === type);
        const matchesCategory = (category === 'ALL' || category === '' || !category) ? true : (txCategory === category);

        if (matchesType && matchesCategory) {
          
          if (txType === 'Income') {
            totalIncome += amount;
          } else if (txType === 'Expense') {
            totalExpense += amount;
          }

          // Return signed amount for table styling (positive for Income, negative for Expense)
          const signedAmount = (txType === 'Expense') ? -amount : amount;

          filteredTransactions.push({
            date: Utilities.formatDate(txDate, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd"),
            accountName: accMap[rowAccId] || rowAccId || 'N/A',
            category: txCategory,
            description: row[6] || '',
            type: txType,
            amount: signedAmount
          });
        }
      }
    }
  });

  const closingBalance = openingBalance + totalIncome - totalExpense;

  return {
    openingBalance: openingBalance,
    totalIncome: totalIncome,
    totalExpense: totalExpense,
    closingBalance: closingBalance,
    transactions: filteredTransactions
  };
}