function dateTest(row){
    var myGoogleSheet = SpreadsheetApp.getActiveSpreadsheet();
    var customersheet = myGoogleSheet.getSheetByName(formObject.customer);
    var data = customersheet.getRange(row,1).getValue();
    return Object.prototype.toString.call(data) === '[object Date]';
}

function getAllSheetNames() {
  var out = new Array()   
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();  
  var stop = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EXPIRED ---->").getIndex();
for (var i=0 ; i<stop-1 ; i++) 
    out.push( [ sheets[i].getName() ] )
    return out;
  }

  //Function to find the right date in column (if column does not have valid date, discard, else decide if the date is older or newer than the input, else resort to EXPIRE method)
