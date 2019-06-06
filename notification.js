  
var actvieSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = actvieSpreadSheet.getSheets()[0]; 
  var lastRow = activeSheet.getLastRow();
  var approvedColumn = 11;
  var emailColumn = 2;
  var itemDescriptionColumn = 5;
  var vendorNameColumn = 4;
  var requestAmountColumn = 6;
  var sentColumn = 12;
  var currentRow = 2;


function sentEmail(e){
  
  while(currentRow<=lastRow){
    
    var approved = activeSheet.getRange(currentRow, approvedColumn).getValues();
    var sent = activeSheet.getRange(currentRow, sentColumn).getValues();
    
    if(approved == "Yes" && sent != "Yes"){
      var email = activeSheet.getRange(currentRow, emailColumn).getValues();
      var purchaseItemDescription = activeSheet.getRange(currentRow, itemDescriptionColumn).getValues();
      var vendorName = activeSheet.getRange(currentRow, vendorNameColumn).getValues();
      var requestAmountCAD = activeSheet.getRange(currentRow, requestAmountColumn).getValues();
    
      var emailSubject = "Purchase Request " + purchaseItemDescription + " by " + email ; 
      var emailBody = "Your purchase request for " + purchaseItemDescription + " has been approved"; 
      var emailBodyForBilling = "The purchase request by " + email + " for " + purchaseItemDescription  +"  sold by " + vendorName + " with the amount of $" + requestAmountCAD  + " has been approved."; 
    
      MailApp.sendEmail(email, emailSubject, emailBody);  
      MailApp.sendEmail("billing@acsea.ca", emailSubject, emailBodyForBilling );
      activeSheet.getRange(currentRow, sentColumn).setValue("Yes");
      currentRow = currentRow + 1;
  
    }else if(approved == "No" && sent != "Yes"){
      var email = activeSheet.getRange(currentRow, emailColumn).getValues();
      var purchaseItemDescription = activeSheet.getRange(currentRow, itemDescriptionColumn).getValues();
      var vendorName = activeSheet.getRange(currentRow, vendorNameColumn).getValues();
      var requestAmountCAD = activeSheet.getRange(currentRow, requestAmountColumn).getValues();
    
      //var emailSubject = "Purchase Request Denied " + purchaseItemDescription + " by " + email ; 
      var emailSubject = "Purchase Request Denied";
      var emailBody = "Your purchase request for " + purchaseItemDescription + " has been denied. Please see Charlie to discuss the details"; 
      var emailBodyForBilling = "The purchase request by " + email + " for " + purchaseItemDescription  +"  sold by " + vendorName + " with the amount of $" + requestAmountCAD  + " has been denied.";  
      MailApp.sendEmail(email, emailSubject, emailBody);
      MailApp.sendEmail("billing@acsea.ca", emailSubject, emailBodyForBilling);
      activeSheet.getRange(currentRow, sentColumn).setValue("Yes");
      currentRow = currentRow + 1;
      
    }
    else {
      currentRow = currentRow + 1;
    }
  } 
}


function onFormSubmit(e) {
  var values = e.namedValues;
  var emailBody ="" ;
  for (key in values){
    var label = key;
    var data = values[key];
    emailBody += label + ": " + data + "\n";
  };
  emailBody += "https://docs.google.com/spreadsheets/d/1ZZV10O2gulafErLXJwOo17ey4Q5P1Bfm4wI1UvK4prs/edit?usp=sharing"

  var emailSubject = "A purchase request by " + values["Email Address"] + " with the amount of $" + values["Request Amount Before Tax"];
 
  MailApp.sendEmail('cwu@acsea.ca', emailSubject, emailBody)
}
             
