function sendEmails() {
  /*
  Sends emails to contact list (Emails!A$2:D) based on spreadsheet template (Templates!$A$1)
  */
  
  // Setup Sheet Variables
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var templateSheet = activeSpreadsheet.getSheetByName('Template');
  var emailsSheet = activeSpreadsheet.getSheetByName('Emails');
  
  emailsSheet.activate();
  
  var ss = activeSpreadsheet.getActiveSheet();
  
  var lastRow = ss.getLastRow();
  var templateText = templateSheet.getRange(1, 1).getValue(); // String containg desired template
  var senderName = 'Vivek Jha';
    
  var quotaLeft = MailApp.getRemainingDailyQuota(); // Can only send 100 emails per day
  Logger.log(quotaLeft);
  if (quotaLeft < lastRow - 1) {
    
    throwQuotaError(quotaLeft, lastRow)
    
  } else {
    
    for (var i = 2; i <= lastRow; i++) {
    
      var emailAddress = ss.getRange(i, 1).getValue();
      var currentName = ss.getRange(i, 2).getValue();
      var currentClassTitle = ss.getRange(i, 3).getValue();
      var currentDate = ss.getRange(i, 4).getDisplayValue();
      
      // JS object with template fields (update accordingly)
      var info = {
      
        name: currentName,
        title: currentClassTitle,
        date: currentDate
        
      };
      
      var subjectLine = 'Reminder: ' + currentClassTitle + ' Upcoming Class';
      var customEmailMsg = populateTemplate(info, templateText);
      Logger.log(customEmailMsg);

      // JS object with email fields (update accordingly)      
      var emailObj = {
        
        to: emailAddress,
        subject: subjectLine,
        body: customEmailMsg,
        name: senderName
        
      };
      
      MailApp.sendEmail(emailObj);
      
    }
    
  }
  
}