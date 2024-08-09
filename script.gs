// Call during first use to set up properties
function setEnvironmentalVariables() {
    var documentProperties = PropertiesService.getDocumentProperties();
    var newProperties = {UPS_CLIENT_ID: '', UPS_CLIENT_SECRET: ''};
    documentProperties.setProperties(newProperties);
}

async function generateUPSToken() {

    const formData = {
        grant_type: 'client_credentials'
    }

    var url = "https://onlinetools.ups.com/security/v1/oauth/token";

    var documentProperties = PropertiesService.getDocumentProperties();
    var auth = 'Basic ' + Utilities.base64EncodeWebSafe(`${documentProperties.getProperty('UPS_CLIENT_ID')}:${documentProperties.getProperty('UPS_CLIENT_SECRET')}`);

    var options = {
      "method": "post",
      "payload": formData,
      "headers": {
          "Content-Type": "application/x-www-form-urlencoded",
          "Authorization": auth,
          },
      };

    const response = await UrlFetchApp.fetch(url, options).getContentText();
    var data = JSON.parse(response)
    console.log(data["access_token"]);
}

function createUPSReturnLabel(form_data) {
  var userEmail = form_data["user_email"]
  var packageType = form_data["return_type"]
  var userData = parseSheetForEmail(userEmail)
  console.log(userData)
}

function parseSheetForEmail(email) {
  var currentSheet = SpreadsheetApp.getActiveSheet();
  var textFinder = currentSheet.createTextFinder(email)
  var foundRange = textFinder.findNext();

  if (foundRange) {
    var rowNumber = foundRange.getRow();
    var userData = currentSheet.getRange(rowNumber, 1, 1, 8).getValues().flat()
    return userData
  } else {
      SpreadsheetApp.getUi().alert(`${email} was not found within the sheet.\n Please enter a different email address and try again.`);
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('UPS Return Tool')
      .addItem('Show sidebar', 'showSidebar')
      .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('UPS Return Tool');
  SpreadsheetApp.getUi()
      .showSidebar(html);

}