function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Team Budget Logger')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Helper: Get Spreadsheet dynamically from Script Properties or active file
function getSpreadsheet() {
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

// Helper utility to manually set Spreadsheet ID
function setSpreadsheetId(id) {
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', id);
}

// GET: Fetch budgets, exchange rates, and user transactions
function getUserBudgetData() {
  var activeUserEmail = Session.getActiveUser().getEmail();
  var targetEmail = activeUserEmail.toLowerCase().trim();
  var ss = getSpreadsheet();

  // 1. Fetch Exchange Rates
  var settingsSheet = ss.getSheetByName("Settings");
  var rates = { "ZAR": 1.0, "BWP": 1.15 }; 
  if (settingsSheet) {
    var settingsData = settingsSheet.getDataRange().getValues();
    for (var r = 1; r < settingsData.length; r++) {
      if (settingsData[r][0]) {
        rates[settingsData[r][0].toString().toUpperCase().trim()] = parseFloat(settingsData[r][1]) || 1.0;
      }
    }
  }

  // 2. Fetch User Transactions first (to sum spending per category)
  var expenseSheet = ss.getSheetByName("Expense_Log");
  var userTransactions = [];
  var categorySpentMap = {}; // Tracks sum of ZAR spent per category

  if (expenseSheet) {
    var expenseData = expenseSheet.getDataRange().getValues();
    for (var j = expenseData.length - 1; j >= 1; j--) {
      if (!expenseData[j][0]) continue;
      var expEmail = expenseData[j][0].toString().toLowerCase().trim();
      
      if (expEmail === targetEmail) {
        var rawDate = expenseData[j][2];
        var formattedDate = rawDate instanceof Date ? rawDate.toLocaleDateString("en-ZA") : rawDate.toString();
        var cat = expenseData[j][3];
        var priceZar = parseFloat(expenseData[j][7]) || 0;

        // Sum spending per category
        categorySpentMap[cat] = (categorySpentMap[cat] || 0) + priceZar;

        userTransactions.push({
          rowIndex: j + 1,
          rowId: expenseData[j][1],
          date: formattedDate,
          category: cat,
          price: parseFloat(expenseData[j][4]) || 0,
          reason: expenseData[j][5] || '',
          currency: expenseData[j][6] || 'ZAR',
          priceZar: priceZar
        });
      }
    }
  }

  // 3. Fetch Budgets & dynamically compute remaining balances
  var budgetSheet = ss.getSheetByName("Budget_Master");
  var budgetData = budgetSheet.getDataRange().getValues();
  var userBudgets = {};
  
  for (var i = 1; i < budgetData.length; i++) {
    if (!budgetData[i][0]) continue;
    var rowEmail = budgetData[i][0].toString().toLowerCase().trim();
    
    if (rowEmail === targetEmail) {
      var category = budgetData[i][1];                    
      var allocated = parseFloat(budgetData[i][2]) || 0;  
      var spent = categorySpentMap[category] || 0;
      var remaining = allocated - spent;            

      if (category) {
        userBudgets[category] = { allocated: allocated, remaining: remaining, spent: spent };
      }
    }
  }

  return {
    email: activeUserEmail,
    budgets: userBudgets,
    transactions: userTransactions,
    rates: rates
  };
}

// POST: Log a new transaction (8-column format with generated Row_ID)
function logExpense(category, price, reason, currency) {
  var activeUserEmail = Session.getActiveUser().getEmail();
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Expense_Log");
  if (!sheet) throw new Error("Sheet 'Expense_Log' not found.");
  
  var settingsSheet = ss.getSheetByName("Settings");
  var rate = 1.0;
  if (currency === 'BWP' && settingsSheet) {
    rate = parseFloat(settingsSheet.getRange("B3").getValue()) || 1.15;
  }

  var priceZar = price * rate;
  var today = new Date().toLocaleDateString("en-ZA");
  var rowId = Utilities.getUuid();
  
  // Appends: User Email | Row_ID | Date | Category | Price | Reason | Currency | Amount(ZAR)
  sheet.appendRow([activeUserEmail, rowId, today, category, price, reason, currency, priceZar]);
  return "Success";
}
// POST: Add a new category to Budget_Master
function addCategory(categoryName, budgetAmount) {
  var activeUserEmail = Session.getActiveUser().getEmail();
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Budget_Master");
  if (!sheet) throw new Error("Sheet 'Budget_Master' not found.");

  var name = categoryName.trim();
  var allocated = parseFloat(budgetAmount) || 0;

  // Append 3 columns: User Email | Category | Budget Allocated
  sheet.appendRow([activeUserEmail, name, allocated]);
  return "Success";
}

// PUT: Update existing transaction by row index
function updateExpense(rowIndex, category, price, reason, currency) {
  var activeUserEmail = Session.getActiveUser().getEmail().toLowerCase().trim();
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Expense_Log");
  
  var rowEmail = sheet.getRange(rowIndex, 1).getValue().toString().toLowerCase().trim();
  if (rowEmail !== activeUserEmail) throw new Error("Unauthorized update request.");

  var settingsSheet = ss.getSheetByName("Settings");
  var rate = 1.0;
  if (currency === 'BWP' && settingsSheet) {
    rate = parseFloat(settingsSheet.getRange("B3").getValue()) || 1.15;
  }
  var priceZar = price * rate;

  // Updates Columns D to H: Category | Price | Reason | Currency | Amount(ZAR)
  sheet.getRange(rowIndex, 4, 1, 5).setValues([[category, price, reason, currency, priceZar]]);
  return "Success";
}

// DELETE: Delete transaction by row index
function deleteExpense(rowIndex) {
  var activeUserEmail = Session.getActiveUser().getEmail().toLowerCase().trim();
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Expense_Log");
  
  var rowEmail = sheet.getRange(rowIndex, 1).getValue().toString().toLowerCase().trim();
  if (rowEmail !== activeUserEmail) throw new Error("Unauthorized delete request.");

  sheet.deleteRow(rowIndex);
  return "Success";
}