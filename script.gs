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
    return data["access_token"];
}

async function createUPSReturnLabel(form_data) {
  var userEmail = form_data["user_email"]
  var packageType = form_data["return_type"]
  var userData = parseSheetForEmail(userEmail)

  // Parameters require "v2403" as version as per https://developer.ups.com/api/reference?loc=en_US#operation/Shipment
  const version = 'v2403';

  const token = generateUPSToken()

  var url = `https://onlinetools.ups.com/api/shipments/${version}/ship`;

  var options = {
    "method": "post",
    "payload": formData,
    "headers": {
        "Content-Type": 'application/json',
        "Authorization": 'Bearer ' + token,
        "transactionSrc": 'testing',
        "transId": '1234'
        },
    "body": JSON.stringify({
    ShipmentRequest: {
      Request: {
        RequestOption: 'nonvalidate',
      },
      Shipment: {
        // 8 =  UPS Electronic Return Label (ERL)
        ReturnService: '8',
        Shipper: {
          Name: 'Redfin',
          AttentionName: 'Redfin IT',
          //ADD AS A PROPERTY
          ShipperNumber: '',
          Address: {
            AddressLine: ['1099 Stewart St #600'],
            City: 'Seattle',
            StateProvinceCode: 'WA',
            PostalCode: '98101',
            CountryCode: 'US'
          }
        },
        ShipTo: {
          Name: `${userData[0]}`,
          EMailAddress: `${userData[1]}`,
          Address: {
            AddressLine: [`${userData[3]} + ${userData[4]}`],
            City: `${userData[5]}`,
            StateProvinceCode: `${userData[6]}`,
            PostalCode: `${userData[7]}`,
            CountryCode: 'US'
          },
        },
        Service: {
          // 03 = Ground
          Code: '03',
        },
        Package: {
          Description: `TICKET NUMBER HERE`,
          Packaging: {
            // 02 = Customer Supplied Packaged
            Code: '02',
          },
          PackageWeight: {
            UnitOfMeasurement: {
              Code: 'LBS',
              Description: 'Pounds'
            },
            //FORM DATA HERE
            Weight: '5'
          }
        }
      },
    }
  })
}
  const response = await UrlFetchApp.fetch(url, options).getContentText();
  var data = JSON.parse(response)
  console.log(data["access_token"]);

};

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