// Call during first use to set up properties
function setEnvironmentalVariables() {
    var documentProperties = PropertiesService.getDocumentProperties();
    var newProperties = {UPS_CLIENT_ID: '', UPS_CLIENT_SECRET: '', UPS_ACCOUNT_NUMBER: ''};
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
    try {
      const response = await UrlFetchApp.fetch(url, options);
      if (response.getResponseCode() === 200) {
        var content = response.getContentText();
        var data = JSON.parse(content)
        return data["access_token"]
      } else {
        SpreadsheetApp.getUi().alert('Error authenticating UPS credentials');
      }
    } catch (error) {
      SpreadsheetApp.getUi().alert('Error authenticating UPS credentials \n Error: \n' + error);
}
}

async function createUPSReturnLabel(form_data) {
  var documentProperties = PropertiesService.getDocumentProperties();
  var userEmail = form_data["user_email"]
  var returnType = form_data["return_type"]
  var userData = parseSheetForEmail(userEmail)

  let ticketNumber;
    if (form_data["ticket_number"]) {
      ticketNumber = form_data["ticket_number"]
    } else {
      ticketNumber = 'n/a'
    }

  let postalCode = postalCodeValidation(userData[7])

  // Parameters require "v2403" as version as per https://developer.ups.com/api/reference?loc=en_US#operation/Shipment
  const version = 'v2403';

  const token = await generateUPSToken()
  var auth = 'Bearer ' + token
  var url = `https://onlinetools.ups.com/api/shipments/${version}/ship`;

  var options = {
    "method": "post",
    "headers": {
        "Content-Type": 'application/json',
        transactionSrc: 'testing',
        transId: '1234',
        Authorization: auth
        },
    "payload": JSON.stringify({
    ShipmentRequest: {
      LabelSpecification: {
        HTTPUserAgent: 'Mozilla/4.5',
        LabelImageFormat: {
          Code: 'EPL',
        },
        LabelStockSize: {
          Height: '6',
          Width: '4'
        }
      },
      Request: {
        RequestOption: 'nonvalidate',
      },
      Shipment: {
        Description: 'Redfin IT Equipment',
        // 8 =  UPS Electronic Return Label (ERL)
        ReturnService: {
          Code: '8',
        },
        Shipper: {
          Name: 'Redfin',
          AttentionName: 'Redfin IT',
          ShipperNumber: `${documentProperties.getProperty('UPS_ACCOUNT_NUMBER')}`,
          EMailAddress: 'jeff.kleinaitis.redfin@titleforward.com',
          Address: {
            AddressLine: ['1099 Stewart St #600'],
            City: 'Seattle',
            StateProvinceCode: 'WA',
            PostalCode: '98101',
            CountryCode: 'US'
          }
        },
      ShipFrom: {
        Name: `${userData[0]}`,
        EMailAddress: `${userData[1]}`,
        Address: {
          AddressLine: [`${userData[3]} ${userData[4]}`],
          City: `${userData[5]}`,
          StateProvinceCode: `${stateNameToAbbreviation(userData[6])}`,
          PostalCode: postalCode,
          CountryCode: 'US'
        },
      },
      ShipTo: {
        Name: 'Redfin IT',
        Address: {
          AddressLine: ['1099 Stewart St #600'],
          City: 'Seattle',
          StateProvinceCode: 'WA',
          PostalCode: '98101',
          CountryCode: 'US'
        },
      },
      PaymentInformation: {
          ShipmentCharge: {
            // 01 = Transportation
            Type: '01',
            BillShipper: {
              AccountNumber: `${documentProperties.getProperty('UPS_ACCOUNT_NUMBER')}`,
            }
          }
        },
      Service: {
        // 03 = Ground
        Code: '03',
      },
      Package: {
        ReferenceNumber: {
          Value: `${ticketNumber}`,
        },
        Description: 'Equipment',
        Packaging: {
          // 02 = Customer Supplied Packaged
          Code: '02',
        },
        PackageWeight: {
          UnitOfMeasurement: {
            Code: 'LBS',
            Description: 'Pounds'
          },
          Weight: `${returnTypeToWeight(returnType)}`
        }
      },
      ShipmentServiceOptions: {
        LabelDelivery: {
          EMail: {
            EMailAddress: `${userData[1]}`,
          }
        },
      },
      },
    }
  })
}
  try {
    const response = await UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() === 200) {
      var content = response.getContentText();
      var data = JSON.parse(content)
      SpreadsheetApp.getUi().alert(`Return shipping label was successfully sent to ${userData[1]}. Tracking number: ${data["ShipmentResponse"]["ShipmentResults"]["ShipmentIdentificationNumber"]}`);
      showSidebar()
    } else {
      SpreadsheetApp.getUi().alert('Creation of return shipping label was not successful.');
      showSidebar()
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert('Creation of return shipping label was not successful.\n Error: \n' + error);
}
};

function parseSheetForEmail(email) {
  var targetSheet = SpreadsheetApp.getActive().getSheetByName('rplSelect');;
  var textFinder = targetSheet.createTextFinder(email)
  var foundRange = textFinder.findNext();

  if (foundRange) {
    var rowNumber = foundRange.getRow();
    var userData = targetSheet.getRange(rowNumber, 1, 1, 8).getValues().flat()
    return userData
  } else {
      SpreadsheetApp.getUi().alert(`${email} was not found within the sheet.\n Please enter a different email address and try again.`);
      showSidebar()
  }
}

function postalCodeValidation(postalCode) {
    return postalCode.toString().padStart(5, '0');
  }

function stateNameToAbbreviation(name) {
	let states = {
		"arizona": "AZ",
		"alabama": "AL",
		"alaska": "AK",
		"arkansas": "AR",
		"california": "CA",
		"colorado": "CO",
		"connecticut": "CT",
		"district of columbia": "DC",
		"delaware": "DE",
		"florida": "FL",
		"georgia": "GA",
		"hawaii": "HI",
		"idaho": "ID",
		"illinois": "IL",
		"indiana": "IN",
		"iowa": "IA",
		"kansas": "KS",
		"kentucky": "KY",
		"louisiana": "LA",
		"maine": "ME",
		"maryland": "MD",
		"massachusetts": "MA",
		"michigan": "MI",
		"minnesota": "MN",
		"mississippi": "MS",
		"missouri": "MO",
		"montana": "MT",
		"nebraska": "NE",
		"nevada": "NV",
		"new hampshire": "NH",
		"new jersey": "NJ",
		"new mexico": "NM",
		"new york": "NY",
		"north carolina": "NC",
		"north dakota": "ND",
		"ohio": "OH",
		"oklahoma": "OK",
		"oregon": "OR",
		"pennsylvania": "PA",
		"rhode island": "RI",
		"south carolina": "SC",
		"south dakota": "SD",
		"tennessee": "TN",
		"texas": "TX",
		"utah": "UT",
		"vermont": "VT",
		"virginia": "VA",
		"washington": "WA",
		"west virginia": "WV",
		"wisconsin": "WI",
		"wyoming": "WY",
	}

  //Trim, remove all non-word characters with the exception of spaces, and convert to lowercase
	let a = name.trim().replace(/[^\w ]/g, "").toLowerCase();
	if(states[a] !== null) {
		return states[a];
	}

	return null;
}

function returnTypeToWeight(return_type) {
  switch(return_type) {
    case `laptop`:
      return '5';
    case `ipad`:
      return '3';
    case `peripherals`:
      return '5';
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('UPS Return Tool')
      .addItem('Create Return Label', 'showSidebar')
      .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('UPS Return Tool');
  SpreadsheetApp.getUi()
      .showSidebar(html);
}