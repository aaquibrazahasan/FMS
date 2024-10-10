function onEdit(e) {
  // Define the sheet name and column number to track
  var ss = SpreadsheetApp.openById('1IGwlbfCJc83lMMxHAvzYJvbbDkoQxp8KNa5E3OKL1_8');
  var ws = ss.getSheetByName('Doers');
  var sheetNameToTrack = "Suggestion & Cordination PMS"; // Name of the specific sheet (tab)
  var columnToTrack = 6; // The column number you want to track (2 = Column B)
  var minRowToTrack = 3; // The row number must be greater than 3
  
  // Get the active sheet, range, column, and row from the event
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var editedColumn = range.getColumn();
  var editedRow = range.getRow();
  var newValue = range.getValue();

if(editedColumn ===10 && editedRow > 3  && e.source.getActiveSheet().getName()==="Suggestion & Cordination PMS" ){
 
 e.source.getActiveSheet().getRange(editedRow,9).setValue(new Date()); } 

if(editedColumn ===13 && editedRow > 3  && e.source.getActiveSheet().getName()==="Suggestion & Cordination PMS" ){
 
 e.source.getActiveSheet().getRange(editedRow,12).setValue(new Date()); } 

if(editedColumn ===17 && editedRow > 3  && e.source.getActiveSheet().getName()==="Suggestion & Cordination PMS" ){
 
 e.source.getActiveSheet().getRange(editedRow,16).setValue(new Date()); } 

if(editedColumn ===23 && editedRow > 3  && e.source.getActiveSheet().getName()==="Suggestion & Cordination PMS" ){
 
 e.source.getActiveSheet().getRange(roeditedRoww,22).setValue(new Date()); } 

if(editedColumn ===28 && editedRow > 3  && e.source.getActiveSheet().getName()==="Suggestion & Cordination PMS" ){
 
 e.source.getActiveSheet().getRange(editedRow,27).setValue(new Date()); } 

if(editedColumn ===32 && editedRow > 3  && e.source.getActiveSheet().getName()==="Suggestion & Cordination PMS" ){
 
 e.source.getActiveSheet().getRange(editedRow,31).setValue(new Date()); } 

if(editedColumn ===36 && editedRow > 3  && e.source.getActiveSheet().getName()==="Suggestion & Cordination PMS" ){
 
 e.source.getActiveSheet().getRange(editedRow,35).setValue(new Date()); } 

if(editedColumn ===40 && editedRow > 3  && e.source.getActiveSheet().getName()==="Suggestion & Cordination PMS" ){
 
 e.source.getActiveSheet().getRange(editedRow,39).setValue(new Date()); } 

  
  // Check if the edit occurred on the correct sheet, column, and row
  if (sheet.getName() === sheetNameToTrack && editedColumn === columnToTrack && editedRow > minRowToTrack) {
    
    // Set up recipient details and dynamic values
    var recipientEmail = ws.getRange("F3").getValue(); // Replace with actual email
    var recipientName = ws.getRange("F2").getValue(); // Replace with dynamic or actual name if needed
    var columnChanged = columnToTrack;
    var row = editedRow;
    var serial_Number = sheet.getRange(editedRow,1).getValue();
    var listOfServices = sheet.getRange(editedRow,2).getValue();
    var suggestion_Coordination = sheet.getRange(editedRow,3).getValue();
    var options_Link = sheet.getRange(editedRow,5).getValue();
  
    
    // Load the HTML email template and pass dynamic values to it
    var template = HtmlService.createTemplateFromFile('index');
    template.recipientName = recipientName;
    template.columnChanged = columnChanged;
    template.row = row;
    template.serial_Number = serial_Number;
    template.listOfServices = listOfServices;
    template.suggestion_Coordination = suggestion_Coordination;
    template.options_Link = options_Link;
    template.newValue = newValue;
    
    
    // Convert the template to HTML
    var htmlBody = template.evaluate().getContent();
    
    // Send the email with the dynamic HTML body
    MailApp.sendEmail({
      to: recipientEmail,
      subject: "Update in Suggestion & Coordination PMS",
      htmlBody: htmlBody
      
    });
  }
}
