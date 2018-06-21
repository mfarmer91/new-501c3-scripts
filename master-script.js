var ui = SpreadsheetApp.getUi();
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getActiveSheet();
var rangeData = sheet.getDataRange();
var lastColumn = rangeData.getLastColumn();
var lastRow = rangeData.getLastRow();
var searchRange = sheet.getRange(2,2, lastRow-1, lastColumn-1);

function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('501c3 Supporter List Hacks')
        .addItem('Step 1 - Highight Rows', 'highlightRows')
        //.addItem('Delete rows', 'readRows')
        .addItem('Step 2 - Clear rows', 'clearRowsThree')
        .addItem('Step 3 - Delete B and H columns', 'deleteBHColumns')
        .addItem('Step 4 - Get sums', 'sumValues')
        .addItem('Step 5 - Add currency formatting', 'formatCurrency')

        .addToUi();
}

function deleteBHColumns() {
  sheet.deleteColumn(2);
  sheet.deleteColumn(7);
};

function formatCurrency() {
  for ( i = 1; i < lastRow; i++ ){
    var oldValue1 = sheet.getRange(i+1, 7).getValue();
    var oldValue2 = sheet.getRange(i+1, 8).getValue();
    sheet.getRange(i+1, 7).setValue("$" + oldValue1);
    sheet.getRange(i+1, 8).setValue("$" + oldValue2);
  }
};

function sumValues() {
  var sumRowSeven = 0;
  var sumRowEight = 0;
  for(var i = 1; i < lastRow; i++){
    sumRowSeven = sumRowSeven + sheet.getRange(i+1, 7).getValue();
    sumRowEight = sumRowEight + sheet.getRange(i+1, 8).getValue();
  }
  sheet.getRange(lastRow + 1, 7).setValue(sumRowSeven).setFontWeight("bold");
  sheet.getRange(lastRow + 1, 8).setValue(sumRowEight).setFontWeight("bold");
  sheet.getRange(lastRow + 2, 7).setValue("Campaign Total").setFontWeight("bold");
  sheet.getRange(lastRow + 2, 8).setValue(sumRowSeven + sumRowEight).setFontWeight("bold");
}


function highlightRows() {
  // Get array of values in the search Range
  var rangeValues = searchRange.getValues();
  // Loop through array and if condition met, add relevant
  // background color.
  for ( i = 0; i < lastColumn - 1; i++){
    for ( j = 0 ; j < lastRow - 1; j++){
      if(rangeValues[j][i].toUpperCase().indexOf("AUSTRIA") > -1
         || rangeValues[j][i].toUpperCase().indexOf("BELGIUM") > -1
         || rangeValues[j][i].toUpperCase().indexOf("BULGARIA") > -1
         || rangeValues[j][i].toUpperCase().indexOf("CROATIA") > -1
         || rangeValues[j][i].toUpperCase().indexOf("HRVATSKA") > -1
         || rangeValues[j][i].toUpperCase().indexOf("CYPRUS") > -1
         || rangeValues[j][i].toUpperCase().indexOf("CZECH REPUBLIC") > -1
         || rangeValues[j][i].toUpperCase().indexOf("DENMARK") > -1
         || rangeValues[j][i].toUpperCase().indexOf("ESTONIA") > -1
         || rangeValues[j][i].toUpperCase().indexOf("FINLAND") > -1
         || rangeValues[j][i].toUpperCase().indexOf("FRANCE") > -1
         || rangeValues[j][i].toUpperCase().indexOf("GERMANY") > -1
         || rangeValues[j][i].toUpperCase().indexOf("GREECE") > -1
         || rangeValues[j][i].toUpperCase().indexOf("GREAT BRITAIN ") > -1
         || rangeValues[j][i].toUpperCase().indexOf("UNITED KINGDOM") > -1
         || rangeValues[j][i].toUpperCase().indexOf("LUXEMBOURG") > -1
         || rangeValues[j][i].toUpperCase().indexOf("MALTA") > -1
         || rangeValues[j][i].toUpperCase().indexOf("NETHERLANDS") > -1
         || rangeValues[j][i].toUpperCase().indexOf("POLAND") > -1
         || rangeValues[j][i].toUpperCase().indexOf("PORTUGAL") > -1
         || rangeValues[j][i].toUpperCase().indexOf("ROMANIA") > -1
         || rangeValues[j][i].toUpperCase().indexOf("SLOVAKIA") > -1
         || rangeValues[j][i].toUpperCase().indexOf("SLOVENIA") > -1
         || rangeValues[j][i].toUpperCase().indexOf("SPAIN") > -1
         || rangeValues[j][i].toUpperCase().indexOf("SWEDEN") > -1
         || rangeValues[j][i].toUpperCase().indexOf("LITHUANIA") > -1
         || rangeValues[j][i].toUpperCase().indexOf("LATVIA") > -1
         || rangeValues[j][i].toUpperCase().indexOf("ITALY") > -1
         || rangeValues[j][i].toUpperCase().indexOf("IRELAND") > -1
         || rangeValues[j][i].toUpperCase().indexOf("HUNGARY") > -1){
        sheet.getRange(j+2,i+2).setBackground("#cc4125"); //address
      }else if (rangeValues[j][i] === 0){
        sheet.getRange(j+2,i+2).setBackground("#e69138");
      };
    };
  };

};

function clearRowsThree() {
  // Get array of values in the search Range
  var rangeValues = searchRange.getValues();
  // Loop through array and if condition met, clear row
  for ( i = 0; i < lastColumn - 1; i++){
    for ( j = 0 ; j < lastRow - 1; j++){
      if(rangeValues[j][i].toUpperCase().indexOf("AUSTRIA") > -1
         || rangeValues[j][i].toUpperCase().indexOf("BELGIUM") > -1
         || rangeValues[j][i].toUpperCase().indexOf("BULGARIA") > -1
         || rangeValues[j][i].toUpperCase().indexOf("CROATIA") > -1
         || rangeValues[j][i].toUpperCase().indexOf("HRVATSKA") > -1
         || rangeValues[j][i].toUpperCase().indexOf("CYPRUS") > -1
         || rangeValues[j][i].toUpperCase().indexOf("CZECH REPUBLIC") > -1
         || rangeValues[j][i].toUpperCase().indexOf("DENMARK") > -1
         || rangeValues[j][i].toUpperCase().indexOf("ESTONIA") > -1
         || rangeValues[j][i].toUpperCase().indexOf("FINLAND") > -1
         || rangeValues[j][i].toUpperCase().indexOf("FRANCE") > -1
         || rangeValues[j][i].toUpperCase().indexOf("GERMANY") > -1
         || rangeValues[j][i].toUpperCase().indexOf("GREECE") > -1
         || rangeValues[j][i].toUpperCase().indexOf("GREAT BRITAIN ") > -1
         || rangeValues[j][i].toUpperCase().indexOf("UNITED KINGDOM") > -1
         || rangeValues[j][i].toUpperCase().indexOf("LUXEMBOURG") > -1
         || rangeValues[j][i].toUpperCase().indexOf("MALTA") > -1
         || rangeValues[j][i].toUpperCase().indexOf("NETHERLANDS") > -1
         || rangeValues[j][i].toUpperCase().indexOf("POLAND") > -1
         || rangeValues[j][i].toUpperCase().indexOf("PORTUGAL") > -1
         || rangeValues[j][i].toUpperCase().indexOf("ROMANIA") > -1
         || rangeValues[j][i].toUpperCase().indexOf("SLOVAKIA") > -1
         || rangeValues[j][i].toUpperCase().indexOf("SLOVENIA") > -1
         || rangeValues[j][i].toUpperCase().indexOf("SPAIN") > -1
         || rangeValues[j][i].toUpperCase().indexOf("SWEDEN") > -1
         || rangeValues[j][i].toUpperCase().indexOf("LITHUANIA") > -1
         || rangeValues[j][i].toUpperCase().indexOf("LATVIA") > -1
         || rangeValues[j][i].toUpperCase().indexOf("ITALY") > -1
         || rangeValues[j][i].toUpperCase().indexOf("IRELAND") > -1
         || rangeValues[j][i].toUpperCase().indexOf("HUNGARY") > -1){
        //sheet.getRange(j+2,i+5).clear(); //donation amount
        //sheet.getRange(j+2,i+4).clear(); //seller profit
        sheet.getRange(j+2,i+3).clear(); //order number
        sheet.getRange(j+2,i+2).clear(); //address
        sheet.getRange(j+2,i+1).clear(); //email
        sheet.getRange(j+2,i).clear(); //middle name
        sheet.getRange(j+2,i-1).clear(); //last name
        sheet.getRange(j+2,i-3).clear(); //order number

      }else if (rangeValues[j][i] === 0){
        sheet.getRange(j+2,i+2).setBackground("#e69138");
      };
    };
  };

};
