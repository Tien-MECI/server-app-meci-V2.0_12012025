require('dotenv').config(); // Thêm dòng này ngay đầu file
const express = require("express");
const { google } = require("googleapis");
const path = require("path");

// === Load credentials từ biến môi trường ===
const credentialsBase64 = process.env.GOOGLE_CREDENTIALS_B64;
if (!credentialsBase64) {
    console.error("GOOGLE_CREDENTIALS_B64 environment variable is missing!");
    process.exit(1);
}

const credentials = JSON.parse(
    Buffer.from(credentialsBase64, "base64").toString("utf-8")
);

// Thay thế toàn bộ \\n bằng \n và trim()
credentials.private_key = credentials.private_key
    .replace(/\\n/g, '\n')
    .trim();
// Sau khi xử lý private key
console.log("Private key starts with:", credentials.private_key.substring(0, 50));
console.log("Private key ends with:", credentials.private_key.slice(-50));
// === Google Auth ===
const scopes = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
];
const auth = new google.auth.JWT(
    credentials.client_email,
    null,
    credentials.private_key,
    scopes
);
const sheets = google.sheets({ version: "v4", auth });

const app = express();
const PORT = process.env.PORT || 3000;
const drive = google.drive({ version: "v3", auth });

// Hàm tải file từ Google Drive về dưới dạng Base64
async function getFileAsBase64(fileId) {
    const res = await drive.files.get(
        { fileId, alt: "media" },
        { responseType: "arraybuffer" }
    );
    const buffer = Buffer.from(res.data, "binary");
    return buffer.toString("base64");
}

// EJS view engine
app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));

// === Spreadsheet ID ===
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
if (!SPREADSHEET_ID) {
    console.error("SPREADSHEET_ID environment variable is missing!");
    process.exit(1);
}

app.get("/", (req, res) => {
    res.send("🚀 Google Sheets API server is running!");
});


// ✅ Endpoint xuất Biên bản giao nhận
app.get("/bbgn", async (req, res) => {
    try {
        console.log("Bắt đầu xuất BBGN...");

        // Lấy mã đơn hàng từ ô B2
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: "Xuat_BB_GN!B2",
        });
        const cellValue = response.data.values ? response.data.values[0][0] : "";

        if (!cellValue) {
            return res.send("⚠️ Ô B2 đang rỗng, chưa có dữ liệu để xuất Biên bản giao nhận.");
        }

        const maDonHang = cellValue;
        console.log(`Mã đơn hàng: ${maDonHang}`);

        // Lấy dữ liệu đơn hàng
        const donHangRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: "Don_hang!A1:CG500903",
        });

        const rows = donHangRes.data.values;
        const data = rows.slice(1);

        const donHang = data.find(row => row[5] === maDonHang);
        if (!donHang) {
            return res.send("❌ Không tìm thấy đơn hàng với mã: " + maDonHang);
        }

        // Lấy chi tiết sản phẩm
        const ctRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: "Don_hang_ct!A1:AD100000",
        });

        const ctRows = ctRes.data.values.slice(1);
        const products = ctRows
            .filter(row => row[1] === maDonHang)
            .map((row, index) => ({
                stt: index + 1,
                tenSanPham: row[9],
                soLuong: row[22],
                donVi: row[23],
                tongSoLuong: row[22],
                ghiChu: row[29] || "",
            }));

        console.log(`Tìm thấy ${products.length} sản phẩm`);

        // ✅ Lấy logo từ Google Drive
        const LOGO_FILE_ID = "1Rwo4pJt222dLTXN9W6knN3A5LwJ5TDIa";
        let logoBase64 = "";

        try {
            const fileMeta = await drive.files.get({
                fileId: LOGO_FILE_ID,
                fields: "mimeType"
            });

            const res = await drive.files.get(
                { fileId: LOGO_FILE_ID, alt: "media" },
                { responseType: "arraybuffer" }
            );

            const buffer = Buffer.from(res.data, "binary");
            logoBase64 = `data:${fileMeta.data.mimeType};base64,${buffer.toString("base64")}`;

            console.log("✅ Logo loaded, mime:", fileMeta.data.mimeType);
        } catch (err) {
            console.error("⚠️ Không lấy được logo:", err.message);
        }



        res.render("bbgn", {
            donHang,
            products,
            logoBase64,
            autoPrint: true,
        });
    } catch (err) {
        console.error("❌ Lỗi xuất BBGN:", JSON.stringify(err, null, 2));
        res.status(500).send("❌ Lỗi khi xuất biên bản giao nhận");
    }

});

// --- Start server ---
app.listen(PORT, () => {
    console.log(`✅ Server is running on port ${PORT}`);
});
app.get("/debug", (req, res) => {
    res.json({
        spreadsheetId: SPREADSHEET_ID,
        clientEmail: credentials.client_email,
        scopes: scopes,
    });
});

app.get("/time", (req, res) => {
    res.send(new Date().toISOString());
});
app.get("/test-auth", async (req, res) => {
    try {
        const token = await auth.getAccessToken();
        res.json({ success: true, token });
    } catch (err) {
        console.error("Auth test failed:", err);
        res.status(500).json({ error: err.message });
    }
});
app.set("views", path.join(__dirname, "views"));
