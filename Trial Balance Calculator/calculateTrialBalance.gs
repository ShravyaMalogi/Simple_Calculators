function onOpen() {
  // Adds a custom menu to the Google Sheets toolbar
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Trial Balance')
    .addItem('Calculate Trial Balance', 'calculateTrialBalance') // Link the menu item to the function
    .addToUi();
}

function calculateTrialBalance() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues(); // Get all data in the sheet
  
  var totalDebits = 0;
  var totalCredits = 0;
  
  // Loop through the rows starting from row 2 to skip headers
  for (var i = 1; i < data.length; i++) { 
    var debitAmount = data[i][1]; // Column B: Debit Amount
    var creditAmount = data[i][3]; // Column D: Credit Amount
    
    // Add to total debits and credits
    if (!isNaN(debitAmount) && debitAmount !== '') {
      totalDebits += debitAmount;
    }
    if (!isNaN(creditAmount) && creditAmount !== '') {
      totalCredits += creditAmount;
    }
  }
  
  // Output the totals to the sheet
  var resultRow = data.length + 1;
  
  sheet.getRange(resultRow, 1).setValue('Total Debits:');
  sheet.getRange(resultRow, 2).setValue(totalDebits);
  
  sheet.getRange(resultRow + 1, 1).setValue('Total Credits:');
  sheet.getRange(resultRow + 1, 2).setValue(totalCredits);
  
  sheet.getRange(resultRow + 2, 1).setValue('Trial Balance Status:');
  
  if (totalDebits === totalCredits) {
    sheet.getRange(resultRow + 2, 2).setValue('Balanced');
  } else {
    sheet.getRange(resultRow + 2, 2).setValue('Unbalanced');
  }
}
