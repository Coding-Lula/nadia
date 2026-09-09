function getCategories() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Categories');
  const data = sheet.getDataRange().getValues();
  const categories = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      categories.push({ id: data[i][0], name: data[i][1], type: data[i][2] });
    }
  }
  return categories;
}

function addCategory(name, type) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Categories');
  const id = generateId('CAT');
  sheet.appendRow([id, name, type]);
  return { status: 'SUCCESS', id };
}

function deleteCategory(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Categories');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.deleteRow(i + 1);
      return { status: 'SUCCESS' };
    }
  }
  return { status: 'ERROR', message: 'Category not found' };
}