function populateTemplate(obj, msg) {
  /*
  Support function to replace template fields with email recipients' features
  */
  
  var messageBody = msg;
  
  for (var [key, val] in obj) {
    
    messageBody = messageBody.replace('{' + key + '}', val);
    
  }

  return messageBody;
  
}


function throwQuotaError(quota_left, last_row) {
  /*
  Displays message box if daily quota for sending emails has been reached
  */

  var title = 'Hit Email Limit';
  var prompt = 'Attempted to send ' + (last_row - 1) + ' emails, but failed\\n(Can only send ' + quota_left + ' more emails today)';
  
  Browser.msgBox(title, prompt, Browser.Buttons.OK);

}