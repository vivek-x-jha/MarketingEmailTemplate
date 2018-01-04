/*
Spreadsheet:
App Academy Sample Marketing Template (11Ax1mLVk7d2KClY0AeXJsIqTMKrlUyUKp2vcbKIRAXw)

Sheet Names:
Contacts (0)
Template (1268994052)

Named Ranges:
sender - Template!B1
subjectLine - Template!B2
templateText - Template!B3
*/
var ss = SpreadsheetApp.getActiveSpreadsheet();
var contactsSheet = ss.getSheetByName('Contacts');
var templateSheet = ss.getSheetByName('Template');

var firstRow = 2;
var lastRow = contactsSheet.getLastRow();
var lastCol = contactsSheet.getLastColumn();


/**
 * Mass email contacts list based on spreadsheet template.
 *
 * @returns {Null}
 */
function sendEmails() {

  // Passed in string ranges must match corresponding named ranges
  var sender = templateSheet.getRange('sender').getValue();
  var subjectLine = templateSheet.getRange('subjectLine').getValue();
  var templateText = templateSheet.getRange('templateText').getValue();
  
  // Can only send 100 emails per day
  var quotaLeft = MailApp.getRemainingDailyQuota();
  
  if (quotaLeft < lastRow - firstRow + 1) {
    
    showHitQuotaMsgBox(quotaLeft, firstRow, lastRow);
    
    } else {
    
    // Loops through each person in contact list
    for (var r = firstRow; r <= lastRow; r++) {
      
      var emailAddress = contactsSheet.getRange(r, 1).getValue();
      var currentFirstName = contactsSheet.getRange(r, 2).getValue();
      var currentPosition = contactsSheet.getRange(r, 4).getValue();
      
      // Object with template fields (update accordingly)
      var fields = {
        
        firstname: currentFirstName,
        position: currentPosition,
        
      };
      
      var customEmailMsg = populateTemplate(fields, templateText);
    
      var emailObj = {
        
        to: emailAddress,
        subject: subjectLine,
        body: customEmailMsg,
        name: sender
        
      };
      
      MailApp.sendEmail(emailObj);
      
    };
    
  };

};


/**
 * Displays message box if daily quota for sending emails has been reached.
 *
 * @param {Number} quota_left   number of daily emails left (0 <= n <= 100)
 * @param {Number} first_row    first row of contacts data
 * @param {Number} last_row     last row of contacts data
 * @returns {Null}
 */
function showHitQuotaMsgBox(quota_left, first_row, last_row) {

  var title = 'Hit Email Limit';
  var prompt = 'Attempted to send ' + (last_row - first_row + 1) + ' emails, but failed.\\n(Can only send ' + quota_left + ' more emails today)';
  
  Browser.msgBox(title, prompt, Browser.Buttons.OK);

};


/**
 * Replaces any fields in template with individual contact's info.
 *
 * @param {Object} fieldsObject   number of daily emails left (<=100)
 * @param {String} templateString    first row of contacts data
 * @returns {String}
 */
function populateTemplate(fieldsObject, templateString) {

  var template = templateString;
  for (var [key, val] in fieldsObject) {
    
    template = template.replace('{' + key + '}', val);
    
  };

  return template;

};


/**
 * Clears contacts list values (keeps the formatting).
 *
 * @returns {Null}
 */
function clearContacts() {

  var contacts = contactsSheet.getRange(firstRow, 1, lastRow, lastCol);
  contacts.clearContent();
  
};