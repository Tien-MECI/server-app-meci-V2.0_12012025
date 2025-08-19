const fs = require('fs');
const { authenticate } = require('@google-cloud/local-auth');
const { google } = require('googleapis');
const path = require('path');
const express = require('express');
const app = express();

// Cấu hình ứng dụng
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
app.use(express.static(path.join(__dirname, 'public')));

// Cấu hình Google Sheets
const SPREADSHEET_ID = '1iBmcwOQGtYoHdS21q1glZTtViG3Ssr6en3tbWsqktdE';
const SHEETS = ['Don_hang', 'Don_hang_ct', 'Xuat_BB_GN'];
const TOKEN_PATH = path.join(__dirname, 'token.json');
const CREDENTIALS_PATH = path.join(__dirname, 'credentials.json');
const LOGO_FILE_ID = '1Rwo4pJt222dLTXN9W6knN3A5LwJ5TDIa';

// Hàm lưu token
function saveToken(token) {
    fs.writeFileSync(TOKEN_PATH, JSON.stringify(token));
}

// Hàm đọc token đã lưu
function loadSavedToken() {
    try {
        return JSON.parse(fs.readFileSync(TOKEN_PATH));
    } catch (err) {
        return null;
    }
}

// Hàm xác thực Google API
async function loadCredentials() {
    const savedToken = loadSavedToken();
    if (savedToken) {
        try {
            const credentials = JSON.parse(fs.readFileSync(CREDENTIALS_PATH));
            const auth = new google.auth.OAuth2(
                credentials.installed.client_id,
                credentials.installed.client_secret,
                credentials.installed.redirect_uris[0]
            );
            auth.setCredentials(savedToken);
            await auth.getAccessToken(); // Kiểm tra token còn hợp lệ
            return auth;
        } catch (err) {
            console.log('Token hết hạn, yêu cầu xác thực lại');
        }
    }

    const auth = await authenticate({
        keyfilePath: CREDENTIALS_PATH,
        scopes: [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive.readonly'
        ],
    });
    saveToken(auth.credentials);
    return auth;
}

// Hàm lấy dữ liệu từ Google Sheets
async function fetchAllSheets() {
    const auth = await loadCredentials();
    const sheets = google.sheets({ version: 'v4', auth });

    const result = {};
    for (const sheetName of SHEETS) {
        try {
            const res = await sheets.spreadsheets.values.get({
                spreadsheetId: SPREADSHEET_ID,
                range: sheetName,
            });
            result[sheetName] = res.data.values || [];
        } catch (err) {
            console.error(`Lỗi khi lấy dữ liệu từ sheet ${sheetName}:`, err.message);
            result[sheetName] = [];
        }
    }
    return result;
}

// Hàm lấy logo dưới dạng base64
async function getLogoBase64() {
    try {
        const auth = await loadCredentials();
        const drive = google.drive({ version: 'v3', auth });

        const res = await drive.files.get({
            fileId: LOGO_FILE_ID,
            alt: 'media'
        }, {
            responseType: 'arraybuffer'
        });

        return Buffer.from(res.data).toString('base64');
    } catch (err) {
        console.error('Không thể lấy logo:', err.message);
        return null;
    }
}

// Route chính - chuyển hướng đến trang in
app.get('/bbgn', async (req, res) => {
    const html = `
    <!DOCTYPE html>
    <html>
    <head>
        <title>In Biên Bản Giao Nhận</title>
        <style>
            body { font-family: Arial, sans-serif; text-align: center; padding: 40px; }
            .container { max-width: 600px; margin: 0 auto; }
            .btn {
                display: inline-block;
                padding: 12px 24px;
                background-color: #4285F4;
                color: white;
                text-decoration: none;
                border-radius: 4px;
                font-size: 16px;
                margin: 10px;
                cursor: pointer;
            }
            .notice { margin: 20px 0; color: #666; }
            .fallback { display: none; margin-top: 30px; }
        </style>
        <script>
            function openPrintWindow() {
                const printWindow = window.open('/bbgn_print?autoprint=1', '_blank');
                
                // Kiểm tra xem popup có bị chặn không
                setTimeout(() => {
                    if (!printWindow || printWindow.closed) {
                        document.getElementById('popupBlocked').style.display = 'block';
                    } else {
                        printWindow.focus();
                    }
                }, 1000);
            }
            
            // Phương án 2: Chuyển hướng trực tiếp nếu popup bị chặn
            function redirectToPrint() {
                window.location.href = '/bbgn_print?autoprint=1';
            }
            
            // Tự động thử mở popup khi trang load
            window.onload = function() {
                openPrintWindow();
                
                // Dự phòng: Nếu sau 2s vẫn chưa mở được thì hiển thị nút
                setTimeout(() => {
                    if (document.getElementById('popupBlocked').style.display === 'block') {
                        document.getElementById('fallbackOptions').style.display = 'block';
                    }
                }, 2000);
            };
        </script>
    </head>
    <body>
        <div class="container">
            <h1>In Biên Bản Giao Nhận</h1>
            
            <div class="notice">Hệ thống đang chuẩn bị tài liệu in...</div>
            
            <div id="popupBlocked" class="fallback">
                <p>⚠️ Trình duyệt đã chặn cửa sổ popup. Vui lòng chọn một trong các phương án sau:</p>
            </div>
            
            <div id="fallbackOptions" class="fallback">
                <a href="/bbgn_print?autoprint=1" target="_blank" class="btn">Mở trong tab mới</a>
                <button onclick="redirectToPrint()" class="btn">Chuyển hướng trực tiếp</button>
                <button onclick="openPrintWindow()" class="btn">Thử lại popup</button>
            </div>
        </div>
    </body>
    </html>
    `;
    res.send(html);
});


// Route trang in ấn
app.get('/bbgn_print', async (req, res) => {
    try {
        console.log('Đang lấy dữ liệu từ Google Sheets...');
        const { Don_hang = [], Don_hang_ct = [], Xuat_BB_GN = [] } = await fetchAllSheets();

        // Kiểm tra dữ liệu hợp lệ
        if (!Xuat_BB_GN || Xuat_BB_GN.length < 2 || !Xuat_BB_GN[1] || Xuat_BB_GN[1].length < 2) {
            throw new Error('Không tìm thấy dữ liệu cần thiết trong sheet Xuat_BB_GN');
        }

        const b2Value = Xuat_BB_GN[1][1];
        if (!b2Value) throw new Error('Không tìm thấy giá trị ô B2');

        console.log('Đang xử lý dữ liệu...');
        const donHangRow = Don_hang.find(row => row && row[5] === b2Value);
        if (!donHangRow) throw new Error(`Không tìm thấy đơn hàng với mã ${b2Value}`);

        const donHangCtRows = Don_hang_ct.filter(row => row && row[1] === b2Value);
        if (donHangCtRows.length === 0) throw new Error('Không tìm thấy chi tiết đơn hàng');

        console.log('Đang tạo logo...');
        const logoBase64 = await getLogoBase64();

        console.log('Đang render template...');
        const shouldAutoPrint = req.query.autoprint === '1';
        res.render('bbgn', {
            donHang: donHangRow,
            products: donHangCtRows.map((row, i) => ({
                stt: i + 1,
                tenSanPham: row[9] || '',
                soLuong: row[22] || '',
                donVi: row[23] || '',
                tongSoLuong: row[21] || '',
                ghiChu: row[24] || ' ',
            })),
            b2Value,
            headersFlags: {
                hasSoLuong: donHangCtRows.some(row => row[22]),
                hasDonVi: donHangCtRows.some(row => row[23]),
                hasTongSoLuong: donHangCtRows.some(row => row[21]),
                hasGhiChu: donHangCtRows.some(row => row[24])
            },
            logoBase64,
            autoPrint: shouldAutoPrint
        });

    } catch (err) {
        console.error('Lỗi:', err);
        res.status(500).send(`
            <div style="text-align: center; padding: 50px;">
                <h2 style="color: #d32f2f;">Lỗi khi tạo biên bản</h2>
                <p>${err.message}</p>
                <a href="/bbgn" style="color: #4285F4;">← Quay lại trang in</a>
            </div>
        `);
    }
});

// Khởi động server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});
