const { google } = require('googleapis');
const { authenticate } = require('@google-cloud/local-auth');
const fs = require('fs');
const path = require('path');

// Đường dẫn lưu token
const TOKEN_PATH = path.join(__dirname, 'token.json');
const CREDENTIALS_PATH = path.join(__dirname, 'credentials.json');

// Hàm lưu token vào file
function saveToken(token) {
    fs.writeFileSync(TOKEN_PATH, JSON.stringify(token));
}

// Hàm kiểm tra token đã lưu
function loadSavedToken() {
    try {
        return JSON.parse(fs.readFileSync(TOKEN_PATH));
    } catch (err) {
        return null; // Không tìm thấy token
    }
}

// Hàm chính để xác thực
async function authorize() {
    // Thử đọc token đã lưu
    const savedToken = loadSavedToken();
    const credentials = JSON.parse(fs.readFileSync(CREDENTIALS_PATH));

    if (savedToken) {
        const auth = new google.auth.OAuth2(
            process.env.CLIENT_ID || credentials.installed.client_id,
            process.env.CLIENT_SECRET || credentials.installed.client_secret,
            process.env.REDIRECT_URIS || credentials.installed.redirect_uris[0]
        );

        // Thiết lập credentials và xử lý sự kiện token
        auth.setCredentials(savedToken);

        // Thêm sự kiện lắng nghe khi token được làm mới
        auth.on('tokens', (tokens) => {
            if (tokens.refresh_token) {
                // Nếu có refresh_token mới thì lưu lại
                const newTokens = {
                    ...savedToken,
                    refresh_token: tokens.refresh_token
                };
                saveToken(newTokens);
            } else if (tokens.access_token) {
                // Cập nhật access_token mới
                const newTokens = {
                    ...savedToken,
                    access_token: tokens.access_token,
                    expiry_date: tokens.expiry_date
                };
                saveToken(newTokens);
            }
        });

        // Kiểm tra xem token còn hợp lệ không
        try {
            await auth.getAccessToken();
            return auth;
        } catch (err) {
            console.log('Token đã hết hạn, yêu cầu xác thực lại');
            // Xóa token cũ nếu không hợp lệ
            fs.unlinkSync(TOKEN_PATH);
        }
    }

    // Nếu không có token hoặc token hết hạn, tạo mới
    const auth = await authenticate({
        keyfilePath: CREDENTIALS_PATH,
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    // Lưu token để dùng lần sau
    saveToken(auth.credentials);

    // Thiết lập sự kiện lắng nghe cho auth mới
    auth.on('tokens', (tokens) => {
        if (tokens.refresh_token) {
            // Lưu lại refresh token để dùng lâu dài
            const newTokens = {
                ...auth.credentials,
                refresh_token: tokens.refresh_token
            };
            saveToken(newTokens);
        } else if (tokens.access_token) {
            // Cập nhật access token mới
            const newTokens = {
                ...auth.credentials,
                access_token: tokens.access_token,
                expiry_date: tokens.expiry_date
            };
            saveToken(newTokens);
        }
    });

    return auth;
}

module.exports = authorize;