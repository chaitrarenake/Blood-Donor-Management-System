// ======================================
// BLOODCONNECT - BLOOD DONOR MANAGEMENT
// BACKEND: GOOGLE APPS SCRIPT
// DATABASE: GOOGLE SHEETS
// ======================================

// YOUR GOOGLE SPREADSHEET ID
const SPREADSHEET_ID = "1bwAeY9IxjUNhyJ2y_n2M90tmJr_ug_8wUuMcM-vXxS0";


// ======================================
// CONNECT TO GOOGLE SPREADSHEET
// ======================================

function getDatabase() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}


// ======================================
// OPEN WEBSITE
// ======================================

function doGet() {

  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('BloodConnect - Blood Donor Management')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');

}


// ======================================
// CREATE DATABASE HEADERS
// ======================================

function setupDatabase() {

  const ss = getDatabase();

  const donorHeaders = [
    'DonorID',
    'FullName',
    'Age',
    'Gender',
    'BloodGroup',
    'Phone',
    'Email',
    'State',
    'District',
    'City',
    'Availability',
    'RegistrationDate'
  ];

  const requestHeaders = [
    'RequestID',
    'PatientName',
    'BloodGroup',
    'Hospital',
    'State',
    'District',
    'City',
    'Units',
    'Phone',
    'RequiredDate',
    'EmergencyLevel',
    'Status',
    'RequestDate'
  ];

  let donors = ss.getSheetByName('Donors');

  if (!donors) {
    donors = ss.insertSheet('Donors');
  }

  donors.getRange(1, 1, 1, donorHeaders.length)
    .setValues([donorHeaders]);

  let requests = ss.getSheetByName('BloodRequests');

  if (!requests) {
    requests = ss.insertSheet('BloodRequests');
  }

  requests.getRange(1, 1, 1, requestHeaders.length)
    .setValues([requestHeaders]);

  return 'BloodConnect database created and connected successfully!';

}


// ======================================
// REGISTER DONOR
// ======================================

function registerDonor(data) {

  const ss = getDatabase();

  const sheet = ss.getSheetByName('Donors');

  if (!sheet) {
    throw new Error('Donors sheet not found. Run setupDatabase first.');
  }

  if (!data.name || !data.bloodGroup ||
      !data.phone || !data.city) {

    throw new Error('Please fill all required fields.');

  }

  const age = Number(data.age);

  if (!Number.isInteger(age) || age < 18 || age > 65) {

    throw new Error('Please enter an age between 18 and 65.');

  }

  if (!/^\d{10}$/.test(String(data.phone))) {

    throw new Error('Enter a valid 10-digit phone number.');

  }

  const validGroups = [
    'A+', 'A-', 'B+', 'B-',
    'AB+', 'AB-', 'O+', 'O-'
  ];

  if (!validGroups.includes(data.bloodGroup)) {
    throw new Error('Invalid blood group.');
  }

  const donorID = 'DON' + new Date().getTime();

  sheet.appendRow([

    donorID,
    data.name,
    age,
    data.gender || '',
    data.bloodGroup,
    String(data.phone),
    data.email || '',
    data.state || '',
    data.district || '',
    data.city,
    'Available',
    new Date()

  ]);

  return 'Donor registered successfully! Your Donor ID is ' + donorID;

}


// ======================================
// SEARCH BLOOD DONORS
// ======================================

function searchDonors(bloodGroup, city) {

  const ss = getDatabase();

  const sheet = ss.getSheetByName('Donors');

  if (!sheet) {
    throw new Error('Donors sheet not found.');
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();

  const results = [];

  for (let i = 0; i < data.length; i++) {

    const row = data[i];

    const donorBlood = String(row[4]).trim().toUpperCase();

    const donorCity = String(row[9]).trim().toLowerCase();

    const availability = String(row[10]).trim().toLowerCase();

    const bloodMatch =
      !bloodGroup ||
      donorBlood === String(bloodGroup).trim().toUpperCase();

    const cityMatch =
      !city ||
      donorCity.includes(String(city).trim().toLowerCase());

    if (bloodMatch && cityMatch && availability === 'available') {

      results.push({

        id: String(row[0]),
        name: String(row[1]),
        age: String(row[2]),
        gender: String(row[3]),
        bloodGroup: String(row[4]),
        city: String(row[9]),
        district: String(row[8])

      });

    }

  }

  return results;

}


// ======================================
// SUBMIT BLOOD REQUEST
// ======================================

function submitBloodRequest(data) {

  const ss = getDatabase();

  const sheet = ss.getSheetByName('BloodRequests');

  if (!sheet) {
    throw new Error('BloodRequests sheet not found.');
  }

  if (!data.name || !data.bloodGroup ||
      !data.phone || !data.hospital) {

    throw new Error('Please fill all required fields.');

  }

  if (!/^\d{10}$/.test(String(data.phone))) {

    throw new Error('Enter a valid 10-digit phone number.');

  }

  const units = Number(data.units);

  if (!Number.isInteger(units) || units < 1 || units > 20) {

    throw new Error('Enter a valid number of units.');

  }

  const validGroups = [
    'A+', 'A-', 'B+', 'B-',
    'AB+', 'AB-', 'O+', 'O-'
  ];

  if (!validGroups.includes(data.bloodGroup)) {
    throw new Error('Invalid blood group.');
  }

  if (!data.requiredDate || !data.emergency) {
    throw new Error('Please enter the required date and emergency level.');
  }

  const requestID = 'REQ' + new Date().getTime();

  sheet.appendRow([

    requestID,
    data.name,
    data.bloodGroup,
    data.hospital,
    data.state || '',
    data.district || '',
    data.city || '',
    units,
    String(data.phone),
    data.requiredDate,
    data.emergency,
    'Pending',
    new Date()

  ]);

  return 'Blood request submitted successfully! Request ID: ' + requestID;

}


// ======================================
// DASHBOARD STATISTICS
// ======================================

function getDashboardStats() {

  const ss = getDatabase();

  const donorSheet = ss.getSheetByName('Donors');

  const requestSheet = ss.getSheetByName('BloodRequests');

  if (!donorSheet || !requestSheet) {
    throw new Error('Database sheets not found. Run setupDatabase first.');
  }

  const donorLastRow = donorSheet.getLastRow();

  const requestLastRow = requestSheet.getLastRow();

  const donors = donorLastRow > 1
    ? donorSheet.getRange(2, 1, donorLastRow - 1, 12).getValues()
    : [];

  const requests = requestLastRow > 1
    ? requestSheet.getRange(2, 1, requestLastRow - 1, 13).getValues()
    : [];

  let availableDonors = 0;

  let pendingRequests = 0;

  for (let i = 0; i < donors.length; i++) {

    if (String(donors[i][10]).trim().toLowerCase() === 'available') {

      availableDonors++;

    }

  }

  for (let i = 0; i < requests.length; i++) {

    if (String(requests[i][11]).trim().toLowerCase() === 'pending') {

      pendingRequests++;

    }

  }

  return {

    totalDonors: donors.length,

    availableDonors: availableDonors,

    totalRequests: requests.length,

    pendingRequests: pendingRequests

  };

}


// ======================================
// TEST DATABASE CONNECTION
// ======================================

function testDatabaseConnection() {

  const ss = getDatabase();

  return {

    success: true,

    message: 'BloodConnect connected successfully!',

    spreadsheetName: ss.getName(),

    spreadsheetURL: ss.getUrl(),

    sheets: ss.getSheets().map(function(sheet) {
      return sheet.getName();
    })

  };

}
