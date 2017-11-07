
function formatMailBody(obj, order) {
  var result = "";

  for (var idx in order) {
    var key = order[idx];
    result += "<h4 style='text-transform: capitalize; margin-bottom: 0'>" + key + "</h4><div>" + obj[key] + "</div>";

  }
  return result;
}

function doPost(e) {

  try {
    Logger.log(e);
    record_data(e);
    
    var mailData = e.parameters;
    

    var dataOrder = JSON.parse(e.parameters.formDataNameOrder);
    
    var sendEmailTo = (typeof TO_ADDRESS !== "undefined") ? TO_ADDRESS : mailData.formGoogleSendEmail;
    
    MailApp.sendEmail({
      to: String(sendEmailTo),
      subject: "Contact form submitted",

      htmlBody: formatMailBody(mailData, dataOrder)
    });

    return ContentService 
          .createTextOutput(
            JSON.stringify({"result":"success",
                            "data": JSON.stringify(e.parameters) }))
          .setMimeType(ContentService.MimeType.JSON);
  } catch(error) {
    Logger.log(error);
    return ContentService
          .createTextOutput(JSON.stringify({"result":"error", "error": e}))
          .setMimeType(ContentService.MimeType.JSON);
  }
}



function record_data(e) {
  Logger.log(JSON.stringify(e)); 
  try {
    var doc     = SpreadsheetApp.getActiveSpreadsheet();
    var sheet   = doc.getSheetByName(e.parameters.formGoogleSheetName); 
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var nextRow = sheet.getLastRow()+1; 
    var row     = [ new Date() ]; 
    for (var i = 1; i < headers.length; i++) {
      if(headers[i].length > 0) {
        row.push(e.parameter[headers[i]]); 
      }
    }
    sheet.getRange(nextRow, 1, 1, row.length).setValues([row]);
  }
  catch(error) {
    Logger.log(e);
  }
  finally {
    return;
  }

}
