// server.js
const express = require('express');
const { google } = require('googleapis');
const authorize = require('./auth-sa');

const app = express();
const PORT = process.env.PORT || 3000;

// ví dụ route xem BBGN
app.get('/bbgn', async (req, res) => {
    try {
        const { orderId } = req.query; // ví dụ lấy tham số từ AppSheet
        const auth = await authorize();
        const sheets = google.sheets({ version: 'v4', auth });

        // ví dụ đọc dữ liệu từ Google Sheet
        // nhớ set env SPREADSHEET_ID trên Render
        const rs = await sheets.spreadsheets.values.get({
            spreadsheetId: process.env.SPREADSHEET_ID,
            range: 'BBGN!A1:G50', // đổi theo sheet/range của bạn
        });

        const rows = rs.data.values || [];

        // TODO: render HTML theo ý bạn (có thể dùng EJS nếu muốn)
        res.send(`
      <html><head><meta charset="utf-8"><title>BBGN</title></head>
      <body style="font-family: system-ui, -apple-system, Arial">
        <h2>BBGN Preview ${orderId ? '(Order: ' + orderId + ')' : ''}</h2>
        <pre>${JSON.stringify(rows, null, 2)}</pre>
      </body></html>
    `);
    } catch (e) {
        console.error(e);
        res.status(500).send('Server error: ' + e.message);
    }
});

app.listen(PORT, () => console.log('Listening on :' + PORT));
