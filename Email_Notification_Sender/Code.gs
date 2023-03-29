// Declare global variables
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getActiveSheet();
var lastRow = sheet.getLastRow();
var lastColumn = sheet.getLastColumn();
var scriptProperties = PropertiesService.getScriptProperties();


// ------------------------ Create User Interface ------------------------

function onOpen() {
  var ui = SpreadsheetApp.getUi(); // Or DocumentApp or SlidesApp or FormApp.
  ui.createMenu('Email System')
    .addItem('Form', 'showFormSidebar')
    .addSeparator()
    .addSubMenu(ui.createMenu('Misc')
      .addItem('Delete', 'deleteData'))
    .addToUi();
}

function showFormSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Form')
    .setTitle('Edit Data')
    .setWidth(300);
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

// ------------------------ Save Sidebar Options ------------------------
function userInput(form) {
  Logger.log('userinput');
  scriptProperties.deleteAllProperties();
  scriptProperties.setProperty('emailSubject', form.emailSubject);
  scriptProperties.setProperty('messageTitle', form.messageTitle);
  scriptProperties.setProperty('messageBody', form.messageBody);
  scriptProperties.setProperty('numToEmail', form.numToEmail);
  scriptProperties.setProperty('userEmailRadio', form.userEmailRadio);
  if (form.userEmailRadio == 'onUserEmailRadio')
    scriptProperties.setProperty('userEmailColumn', form.userEmailColumn);
  scriptProperties.setProperty('emailsString',  makeEmailsString(form));
  scriptProperties.setProperty('newEmailBoxes', addEmailBoxes(form));
  scriptProperties.setProperty('buttonRadio', form.buttonRadio);
  if (form.buttonRadio == 'onButton')
    scriptProperties.setProperty('buttonURL', ss.getUrl());
  // logProperties();
}

// --------------- Returns Array of All Properties  ---------------------
function getAllProperties() {
  var propertiesAndKeys = {}
  var data = scriptProperties.getProperties();
  for (var key in data) {
    propertiesAndKeys[key] = scriptProperties.getProperty(key);
    // Logger.log('Key: %s - %s', key, data[key]);
  }
  return propertiesAndKeys;
}

// Used to turn sidebar emails into a single string for sending
function makeEmailsString(form) {
  var emailString = '';
  var email;
  for (var i = 1; i < parseInt(form.numToEmail) + 1; i++) {
    email = form['email' + i];
    emailString = emailString + email + ',';
  }
  emailString = emailString.substring(0, emailString.length - 1);
  return emailString;
}

// Creates additional html email address boxes for the sidebar
function addEmailBoxes(form) {
  var amount = parseInt(form.numToEmail);
  var html = ''
  for (var i = 1; i < amount + 1; i++) {
    var value = form['email' + i] || '';
    html = html + i + ':<input type="text" value="' + value + '" id="email' + i +
      '" name="email' + i + '"/><br>';
  }
  return html;
}

// Limits the Saved message in the sidebar to 6 seconds
function waitSeconds() {
  Utilities.sleep(6000);
}

// Deletes all properties
function deleteData() {
  scriptProperties.deleteAllProperties();
}

// ----------------- Log Script Properties ------------------ 
function logProperties() {
  Logger.log('Log Properties');
  var scriptProperties = PropertiesService.getScriptProperties();
  var data = scriptProperties.getProperties();
  for (var key in data) {
    Logger.log('Key: %s, Value: %s', key, data[key]);
  }
}

// ----------------- Send Email on Form Submission ------------------
// Once a form, Google Form is submitted to spreadsheet
function onFormSubmission() {
  var properties = getAllProperties();
  var submission = getSubmissionString(properties);
  var allEmails = checkUserEmail(properties);
  // Makes an email template, updates it with saved data
  var htmlEmail = HtmlService.createTemplateFromFile('Email');
  htmlEmail.messageTitle = properties.messageTitle;
  htmlEmail.messageBody = properties.messageBody;
  htmlEmail = htmlEmail.evaluate().append(submission);
  htmlEmail = htmlEmail.getContent();
  emailSend(properties, htmlEmail, allEmails);
}

// Get sheet header row and last row submitted for html email output
function getSubmissionString(properties) {
  var submission = '<div class="box">';
  for (var i = 1; i < lastColumn + 1; i++) {
    submission = submission + '<span id="label">' + sheet.getRange(1, i).getValue() +
      '</span><br><span id="info">' + sheet.getRange(lastRow, i).getValue() + '</span><br>';
  }
  // Check is button is on in sidebar
  if (properties.buttonRadio == 'onButton') {
    submission = submission + '<br><a href="' + properties.buttonURL +
      '" id="button">View</a><br><br>';
  }
  submission = submission + '</div><br><p style="font-size:12px;'+
   ' text-align:center;">Created by <a href=' +
   '"https://kurtkaiser.us/">Kurt Kaiser</a> </p>';
  return submission;
}

// Check if sidebar user email option, add user email from sheet if it is
function checkUserEmail(properties) {
  Logger.log('in user email check');
  if (properties.userEmailRadio == 'onUserEmailRadio' && properties.emailsString) {
    properties.emailsString = properties.emailsString + ',' + 
        sheet.getRange(lastRow, parseInt(properties.userEmailColumn)).getValue();
  }
  return properties.emailsString;
}

// Send the email to the required parties
function emailSend(properties, htmlEmail, allEmails) {
  Logger.log('in emailSend');
  MailApp.sendEmail({
    to: allEmails, 
    subject: properties.emailSubject,
    htmlBody: htmlEmail
  })
}
