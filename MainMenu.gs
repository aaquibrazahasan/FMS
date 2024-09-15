function onOpen() {
  var ofS = 8 //this is office start time. 8 means 8 in the morning. Please search for ofS again as there are two places where you will need to change the working hours.
  var ofE = 18 //this is office end time. 18 means 6 in the evening. Please search for ofE again as there are two places where you will need to change the working hours.
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActive();
  ss.getRangeByName('B1').setFormula('='+ofS+'/24-'+ofE+'/24+1')
  ui.createMenu('BMP Formulas')
      .addItem('Assisted ImportRange', 'importRangeFormula')
      .addSeparator()
      .addSubMenu(ui.createMenu('FMS Formulas')
      .addItem('Set Simple TAT', 'addTAT')
      .addSeparator()
      .addItem('Set Working Hours Formula', 'workingHoursFormula')
      .addSeparator()
      .addItem('Set Actual Time Formula', 'createAutoTime')
      .addSeparator()
      .addItem('Set Yellow CF', 'createCFormula'))
      .addToUi();
}
 

function addTAT() {
 var ss = SpreadsheetApp.getActive();
 var ui = SpreadsheetApp.getUi();
 var formulaCell = ui.prompt('Cell to apply formula on', ui.ButtonSet.OK_CANCEL);

 if (formulaCell.getSelectedButton() == ui.Button.OK) {
   var formulaCell = formulaCell.getResponseText();
 } 
  
  var fromDate = ui.prompt('Date cell range.', ui.ButtonSet.OK_CANCEL);
 
  if (fromDate.getSelectedButton() == ui.Button.OK) {
   var fromDate = fromDate.getResponseText();
 } 
  
 var tatInHours = ui.prompt('TAT cell (example F$5)', ui.ButtonSet.OK_CANCEL);
  if (tatInHours.getSelectedButton() == ui.Button.OK) {
  var tatInHours = tatInHours.getResponseText()
  }
   
ss.getRange(formulaCell).activate();
  ss.getCurrentCell().setFormula('=if(' + fromDate + ',' + fromDate + '+' + tatInHours + ',"")')
}


function workingHoursFormula() {
  var ss = SpreadsheetApp.getActive();
 var ui = SpreadsheetApp.getUi();
 var formulaCell = ui.prompt('Cell to apply formula', ui.ButtonSet.OK_CANCEL);

 if (formulaCell.getSelectedButton() == ui.Button.OK) {
   var formulaCell = formulaCell.getResponseText();
 } 
  
 var fromDate = ui.prompt('Date cell range. For example if the date is in A7, then write A7:A', ui.ButtonSet.OK_CANCEL);
 
  if (fromDate.getSelectedButton() == ui.Button.OK) {
   var fromDate = fromDate.getResponseText();
 } 
  
 var tatInHours = ui.prompt('TAT cell, (example F$5)', ui.ButtonSet.OK_CANCEL);
  if (tatInHours.getSelectedButton() == ui.Button.OK) {
  var tatInHours = tatInHours.getResponseText()
  }
  
 ui.alert("You don't have to drag down this formula, Also change the format of Planned column to Date and time format.",ui.ButtonSet.OK)
 
  var ofS = 8 //this is office start time. 8 means 8 in the morning.
  var ofE = 18 //this is office end time. 18 means 6 in the evening. 

  var newdate = fromDate + '+' + tatInHours
ss.getRange(formulaCell).activate();
  ss.getCurrentCell().setFormula(getWFormula(fromDate,tatInHours,ofS,ofE))
}


// --------------------------------------------------------------------------------------------------------------------------

function getWFormula(originalC, tatInHours,ofS,ofE) {
  return '=arrayformula(if('+originalC+',if(weekday(if(--(hour('+originalC+'+'+tatInHours+')>'+ofS+') * --(hour('+originalC+'+'+tatInHours+')<'+ofE+'),'+originalC+'+'+tatInHours+','+originalC+'+'+tatInHours+'+$B$1))=1,if(--(hour('+originalC+'+'+tatInHours+')>'+ofS+') * --(hour('+originalC+'+'+tatInHours+')<'+ofE+'),'+originalC+'+'+tatInHours+','+originalC+'+'+tatInHours+'+$B$1)+1,if(--(hour('+originalC+'+'+tatInHours+')>'+ofS+') * --(hour('+originalC+'+'+tatInHours+')<'+ofE+'),'+originalC+'+'+tatInHours+','+originalC+'+'+tatInHours+'+$B$1)),""))'
}

// --------------------------------------------------------------------------------------------------------------------------

function createAutoTime() {
  var ss = SpreadsheetApp.getActive();
 var ui = SpreadsheetApp.getUi();
  
 ui.alert('Please put =now() in cell A1');
 var formulaCell = ui.prompt('Cell to apply formula', ui.ButtonSet.OK_CANCEL);

 if (formulaCell.getSelectedButton() == ui.Button.OK) {
   var formulaCell = formulaCell.getResponseText();
 } 
 var nextCell = ui.prompt('Cell next to formula cell', ui.ButtonSet.OK_CANCEL);

 if (nextCell.getSelectedButton() == ui.Button.OK) {
   var nextCell = nextCell.getResponseText();
 }  
  
 ss.getRange(formulaCell).activate();
  ss.getCurrentCell().setFormula('=if('+formulaCell+','+formulaCell+',if('+nextCell+'<>"",$a$1,"")')
}


// --------------------------------------------------------------------------------------------------------------------------

function importRangeFormula() {
  var ss = SpreadsheetApp.getActive();
 var ui = SpreadsheetApp.getUi();
 var formulaCell = ui.prompt('Cell to apply formula on', ui.ButtonSet.OK_CANCEL);

 if (formulaCell.getSelectedButton() == ui.Button.OK) {
   var formulaCell = formulaCell.getResponseText();
 } 
 var donorSheet = ui.prompt('URL of the spreadsheet from where \ndata has to be taken', ui.ButtonSet.OK_CANCEL);

 if (donorSheet.getSelectedButton() == ui.Button.OK) {
   var donorSheet = donorSheet.getResponseText();
 }  
  
 var tabName = ui.prompt('Name of the sheet/tab from where \ndata has to be taken', ui.ButtonSet.OK_CANCEL);

 if (tabName.getSelectedButton() == ui.Button.OK) {
   var tabName = tabName.getResponseText();
 }
  
 var rangeName = ui.prompt('Enter the range from where \ndata has to be taken.', ui.ButtonSet.OK_CANCEL);

 if (rangeName.getSelectedButton() == ui.Button.OK) {
   var rangeName = rangeName.getResponseText();
 }
  
 ss.getRange(formulaCell).activate();
  ss.getCurrentCell().setFormula('=importrange("' + donorSheet + '","' + tabName + '!' + rangeName + '")')
  
  
}
  
 
// --------------------------------------------------------------------------------------------------------------------------

//-------------------------------------------------------------------------------------------------------------------------

function createCFormula() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  
  ui.alert("Please make sure that you have selected the cell where you want to put conditional formatting",ui.ButtonSet.OK)
  
  var plannedCell = ui.prompt('Enter the planned cell', ui.ButtonSet.OK_CANCEL);
  
  if (plannedCell.getSelectedButton() == ui.Button.OK) {
   var plannedCell = plannedCell.getResponseText();
 }
  
 setCF(plannedCell)
    
  
}

// --------------------------------------------------------------------------------------------------------------------------


function setCF(planned) {
  var spreadsheet = SpreadsheetApp.getActive();
  var conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getActiveRange()])
  .whenCellNotEmpty()
  .setBackground('#B7E1CD')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.splice(conditionalFormatRules.length - 1, 1, SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getActiveRange()])
  .whenFormulaSatisfied('=if(G7,if(H7,FALSE,if($A$1>G7,TRUE,FALSE)),FALSE)')
  .setBackground('#B7E1CD')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.splice(conditionalFormatRules.length - 1, 1, SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getActiveRange()])
  .whenFormulaSatisfied('=if('+planned+',if('+spreadsheet.getActiveRange().getA1Notation()+',FALSE,if($A$1>'+planned+',TRUE,FALSE)),FALSE)')
  .setBackground('#FCE8B2')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
};


// This function is used to create a Timestamp based on column-specific changes  

function onEdit(e) {
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
 
  var range= sheet.getActiveCell();
  var row = range.getRow();
  var col = range.getColumn();
  
 

if(col ===8 && row > 2  && e.source.getActiveSheet().getName()==="8th Sept - Guest Check-out" ){
 
 e.source.getActiveSheet().getRange(row,9).setValue(new Date()); } 




}

