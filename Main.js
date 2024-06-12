/**Api Key Setter */
function saveApiKey(apiKey) {
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('API_KEY', apiKey);
}
/**
 * Api Key Getter 
 * returns a key in str format
*/
function getApiKey() {
  var userProperties = PropertiesService.getUserProperties();
  return userProperties.getProperty('API_KEY');
}


/**
 * Function that tests that the user's key is valid by sending a dummy request
 */
function verifyApiKey(apiKey) {
  var apiCall = `https://api.hunter.io/v2/email-finder?domain=reddit.com&first_name=Alexis&last_name=Ohanian&api_key=${apiKey}`;

  try {
    var response = UrlFetchApp.fetch(apiCall);
    var result = JSON.parse(response.getContentText());

    if (result.data.email === "alexis@reddit.com") {
      saveApiKey(apiKey);
      return 'API Key saved successfully!';
    } else {
      return 'Invalid or expired API Key.';
    }
  } catch (e) {
    return 'Invalid or expired API Key.';
  }
}

function ensureDotCom(domain) {
  if (!domain.endsWith('.com')) {
    domain += '.com';
  }
  return domain;
}

/**
 * Function that builds the API call, fetches the response and filters the email. 
 */

function hunterEmailRequest(firstName, lastName, company) { 
  var apiKey = getApiKey();
  if (!apiKey) {
    throw new Error('API Key not found. Please set the API Key first.');
  }
  
  var domain = ensureDotCom(company);
  var apiCall = `https://api.hunter.io/v2/email-finder?domain=${domain}&first_name=${firstName}&last_name=${lastName}&api_key=${apiKey}`;
  
  var response = UrlFetchApp.fetch(apiCall);
  var result = JSON.parse(response.getContentText());
  var email = result.data.email;

  return email;
}

/**
 * Custom function to retrieve first name, last name, and company from specified cells.
 * @param {Range} firstNameRange The cells containing the first names.
 * @param {Range} lastNameRange The cells containing the last names.
 * @param {Range} companyRange The cells containing the companies.
 * @return {string[]} The array of emails.
 * @customfunction
 */
function FindEmails(firstNameRange, lastNameRange, companyRange) {
  
  // Initialize a 2D array to store the results, results are inserted in the column called in order
  var results = [];

  // Tests if only one cell per range has been called or not
  if(typeof(firstNameRange)==='string'){
    try {
      var email = hunterEmailRequest(firstNameRange, lastNameRange, companyRange);
      results.push(email);
    } catch (e) {
      Logger.log('Error for ' + firstName + ' ' + lastName + ' at ' + company + ': ' + e.message);
      results = 'Invalid domain or request failed';
    }
    return results;
  }
  for (var i = 0; i < firstNameRange.length; i++) {
    var rowResult = [];
    for (var j = 0; j < firstNameRange[i].length; j++) {
      var firstName = firstNameRange[i][j];
      var lastName = lastNameRange[i][j];
      var company = companyRange[i][j];
      try {
        var email = hunterEmailRequest(firstName, lastName, company);
        rowResult.push(email);
      } catch (e) {
        rowResult.push('Invalid domain or request failed');
      }
    }
    results.push(rowResult);
  }

  return results;
}

function testGetApiKey() {
  var apiKey = getApiKey();
}
/** 
 * Adds a custom menu to the Google Sheets UI when the spreadsheet is opened.
*/
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Email Finder')
    .addItem('Set API Key', 'showApiKeySidebar')
    .addItem('Open Email Finder', 'showSidebar')
    .addToUi();
}
/**
 * Displays a sidebar for setting the API key.
 */
function showApiKeySidebar() {
  var html = HtmlService.createHtmlOutputFromFile('ApiKey')
      .setTitle('Set API Key')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}
/**
 * Displays the main Email Finder sidebar, and if the API key is not set, it shows the API key sidebar instead.
 */
function showSidebar() {
  var apiKey = getApiKey();
  if (!apiKey) {
    showApiKeySidebar();
    return;
  }
  
  var html = HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Email Finder')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}
