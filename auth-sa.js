// auth-sa.js
const { google } = require('googleapis');

module.exports = async function authorize() {
    const b64 = process.env.GOOGLE_CREDENTIALS_B64;
    if (!b64) throw new Error('Missing GOOGLE_CREDENTIALS_B64');

    const json = Buffer.from(b64, 'base64').toString('utf8');
    const creds = JSON.parse(json);

    // Fix xuống dòng private key nếu cần
    const privateKey = (creds.private_key || '').replace(/\\n/g, '\n');

    const auth = new google.auth.JWT(
        creds.client_email,
        null,
        privateKey,
        ['https://www.googleapis.com/auth/spreadsheets'] // thêm scopes khác nếu cần
    );

    await auth.authorize();
    return auth;
};
