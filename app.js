const { google } = require('googleapis');
const path = require('path');

// Load credentials từ file JSON
const credentials = require('./credentials.json');

// Khởi tạo OAuth2 với service account (JWT)
const scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
];
const auth = new google.auth.JWT(
    credentials.client_email,
    null,
    credentials.private_key,
    scopes
);

// Khởi tạo client cho Sheets và Drive
const sheets = google.sheets({ version: 'v4', auth });
const drive = google.drive({ version: 'v3', auth });
