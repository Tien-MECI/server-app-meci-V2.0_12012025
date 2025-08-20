import dotenv from "dotenv";
dotenv.config();
const express = require("express");
const { google } = require("googleapis");
const path = require("path");
const LOGO_FILE_ID = "1Rwo4pJt222dLTXN9W6knN3A5LwJ5TDIa";

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


import puppeteer from "puppeteer";
// ✅ Endpoint xuất Biên bản giao nhận + tự động tạo PDF
app.get("/bbgn", async (req, res) => {
    try {
        console.log("Bắt đầu xuất BBGN...");

        // 1. Lấy mã đơn hàng từ dòng cuối cột B sheet file_BBGN_ct
        const bbgnRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: "file_BBGN_ct!B:B",
        });
        const colB = bbgnRes.data.values ? bbgnRes.data.values.flat() : [];
        const lastRowWithData = colB.length;
        const maDonHang = colB[lastRowWithData - 1];
        console.log(`Mã đơn hàng: ${maDonHang} (dòng ${lastRowWithData})`);

        if (!maDonHang) {
            return res.send("⚠️ Không tìm thấy dữ liệu ở cột B sheet file_BBGN_ct.");
        }

        // 2. Lấy dữ liệu đơn hàng
        const donHangRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: "Don_hang!A1:CG500903",
        });
        const rows = donHangRes.data.values;
        const data = rows.slice(1);
        const donHang = data.find(row => row[5] === maDonHang);
        if (!donHang) return res.send("❌ Không tìm thấy đơn hàng với mã: " + maDonHang);

        // 3. Lấy chi tiết sản phẩm
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

        // 4. Lấy logo
        const LOGO_FILE_ID = "1Rwo4pJt222dLTXN9W6knN3A5LwJ5TDIa";
        let logoBase64 = "";
        try {
            const fileMeta = await drive.files.get({ fileId: LOGO_FILE_ID, fields: "mimeType" });
            const resFile = await drive.files.get({ fileId: LOGO_FILE_ID, alt: "media" }, { responseType: "arraybuffer" });
            const buffer = Buffer.from(resFile.data, "binary");
            logoBase64 = `data:${fileMeta.data.mimeType};base64,${buffer.toString("base64")}`;
        } catch (err) {
            console.error("⚠️ Không lấy được logo:", err.message);
        }

        // 5. Lấy watermark
        const WATERMARK_FILE_ID = "1fNROb-dRtRl2RCCDCxGPozU3oHMSIkHr";
        let watermarkBase64 = "";
        try {
            const fileMeta = await drive.files.get({ fileId: WATERMARK_FILE_ID, fields: "mimeType" });
            const resFile = await drive.files.get({ fileId: WATERMARK_FILE_ID, alt: "media" }, { responseType: "arraybuffer" });
            const buffer = Buffer.from(resFile.data, "binary");
            watermarkBase64 = `data:${fileMeta.data.mimeType};base64,${buffer.toString("base64")}`;
        } catch (err) {
            console.error("⚠️ Không lấy được watermark:", err.message);
        }

        // 6. Tạo PDF với puppeteer và lưu lên Drive
        const today = new Date();
        const dd = String(today.getDate()).padStart(2, "0");
        const mm = String(today.getMonth() + 1).padStart(2, "0");
        const yyyy = today.getFullYear();
        const hh = String(today.getHours()).padStart(2, "0");
        const mi = String(today.getMinutes()).padStart(2, "0");
        const ss = String(today.getSeconds()).padStart(2, "0");

        const fileName = `BBGN - ${maDonHang} - ${dd}${mm}${yyyy} - ${hh}-${mi}-${ss}.pdf`;

        const browser = await puppeteer.launch({ headless: "new" });
        const page = await browser.newPage();

        // ⚠️ Dùng endpoint riêng hiển thị HTML sạch (không autoPrint) để render PDF
        await page.goto(`https://hsdh-app-cu.onrender.com/bbgn-view?maDonHang=${maDonHang}`, {
            waitUntil: "networkidle0"
        });
        const pdfBuffer = await page.pdf({ format: "A4", printBackground: true });
        await browser.close();

        const folderId = "1CL3JuFprNj1a406XWXTtbQMZmyKxhczW";
        const fileMeta = { name: fileName, parents: [folderId] };
        const media = { mimeType: "application/pdf", body: Buffer.from(pdfBuffer) };

        const pdfFile = await drive.files.create({
            requestBody: fileMeta,
            media,
            fields: "id, name"
        });

        const folderMeta = await drive.files.get({ fileId: folderId, fields: "name" });
        const pathToFile = `${folderMeta.data.name}/${pdfFile.data.name}`;

        // Ghi đường dẫn vào cột D dòng tương ứng
        await sheets.spreadsheets.values.update({
            spreadsheetId: SPREADSHEET_ID,
            range: `file_BBGN_ct!D${lastRowWithData}`,
            valueInputOption: "RAW",
            requestBody: { values: [[pathToFile]] }
        });

        // 7. Render bbgn cho người dùng + autoPrint
        res.render("bbgn", {
            donHang,
            products,
            logoBase64,
            watermarkBase64,
            autoPrint: true,
            maDonHang
        });
    } catch (err) {
        console.error("❌ Lỗi xuất BBGN:", err);
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




app.get("/test-logo", async (req, res) => {
    try {
        const fileMeta = await drive.files.get({
            fileId: LOGO_FILE_ID,
            fields: "mimeType"
        });

        const response = await drive.files.get(
            { fileId: LOGO_FILE_ID, alt: "media" },
            { responseType: "arraybuffer" }
        );

        const buffer = Buffer.from(response.data, "binary");
        const base64 = `data:${fileMeta.data.mimeType};base64,${buffer.toString("base64")}`;

        res.send(`<img src="${base64}" style="max-height:100px;">`);
    } catch (err) {
        res.send("Lỗi lấy logo: " + err.message);
    }
});

