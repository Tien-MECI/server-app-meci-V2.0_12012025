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

// Fix xuống dòng cho private_key
if (credentials.private_key) {
    credentials.private_key = credentials.private_key.replace(/\\n/g, "\n");
} else {
    console.error("Private key is missing in credentials!");
    process.exit(1);
}

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

// --- Endpoint đọc sheet theo tên ---
app.get("/sheet/:name", async (req, res) => {
    const sheetName = req.params.name;
    try {
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: `${sheetName}!A:Z`,
        });
        res.json(response.data);
    } catch (err) {
        console.error("❌ Lỗi Google Sheets:", err.errors || err.message || err);
        res.status(500).send(`Error reading sheet "${sheetName}"`);
    }
});

// --- Endpoint đọc toàn bộ sheet đầu tiên ---
app.get("/sheet-all", async (req, res) => {
    try {
        const meta = await sheets.spreadsheets.get({
            spreadsheetId: SPREADSHEET_ID,
        });

        const firstSheet = meta.data.sheets[0].properties.title;
        console.log("📄 Sheet đầu tiên:", firstSheet);

        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: `${firstSheet}!A:Z`,
        });
        res.json(response.data);
    } catch (err) {
        console.error("❌ Lỗi Google Sheets:", err.errors || err.message || err);
        res.status(500).send("Error reading first sheet");
    }
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

        const logoBase64 = ""; // có thể nhúng logo

        res.render("bbgn", {
            donHang,
            products,
            logoBase64,
            autoPrint: false,
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
