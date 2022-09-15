//CREATE CUSTOM MENU
function onOpen() { 
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("SOS")
    .addItem("Update Hours","updateHours")
    .addItem("New Customer","newCustomer")
    .addToUi();
}

//OPEN THE FORM IN SIDEBAR 
function updateHours() { 
  var form = HtmlService.createTemplateFromFile('HoursUpdateIndex').evaluate().setTitle('Engagement Details');
  SpreadsheetApp.getUi().showSidebar(form);
}


//OPEN THE FORM IN MODAL DIALOG
function newCustomer() {
  var form = HtmlService.createTemplateFromFile('NewCustomerIndex').evaluate().setTitle('New Customer');
  SpreadsheetApp.getUi().showSidebar(form);
}

//INPUT DATA TO THE SHEET
function inputData(formObject) {
    var ui = SpreadsheetApp.getUi();
    var myGoogleSheet = SpreadsheetApp.getActiveSpreadsheet();
    var customersheet = myGoogleSheet.getSheetByName((formObject.customer).toString());
    var lastRow = customersheet.getLastRow();
    var inDay = Number(new Date(formObject.date));
  
  //This section goes through all the rows in the first column looking for dates. If it finds one, it returns that information via Boolean true. Otherwise it returns false.
    //The false value will be used to trigger the secondary search, for EXPIRE.
  for ( var row = lastRow; row >= 0; row = row - 1) {
    if (dateTest(row, formObject) == true){
      var sheetday = Number(customersheet.getRange(row, 1).getValue());
      var isadate = true
      if (sheetday <= inDay) {
        var currRow = row;
        break;
  }
      continue;
    }
    }

  if (isadate != true) {
    for (let i = lastRow; i > 0; i--) {
    var search = customersheet.getRange(i, 2).getValue()
    if (search == "Expire" || search == "EXPIRE" || search == "Expired" || search == "EXPIRED") {
      var currRow = i;
      return currRow;
      }

    }}
//End date check

    //Where to put the data
  var moveRow = customersheet.getRange(currRow + 1, 1, 30, 5);
  moveRow.copyValuesToRange(customersheet, 1, 5, currRow + 2, currRow + 30);
    //code to update customer sheet
  
  customersheet.getRange(currRow + 1, 1).setValue(formObject.date);//Engagement Date
  customersheet.getRange(currRow + 1, 2).setValue(formObject.type);//Engagement Type
  customersheet.getRange(currRow + 1, 3).setValue('-' + formObject.hours);//Hours Used
  customersheet.getRange(currRow, 4).copyTo(customersheet.getRange(currRow + 1, 4));//Hours Remaining
  customersheet.getRange(currRow + 1, 5).setValue(formObject.description);//Summary
ui.alert("Customer " + formObject.customer + " has been updated.")


}

//PROCESS FORM

//INCLUDE HTML PARTS, EG. JAVASCRIPT, CSS, OTHER HTML FILES
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}