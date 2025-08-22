import dotenv from "dotenv";
import express from "express";
import { google } from "googleapis";
import path from "path";
import { fileURLToPath } from "url";
import { dirname } from "path";
import ejs from "ejs";
import fetch from "node-fetch"; // Đảm bảo import fetch

dotenv.config();

// --- __dirname trong ESM ---
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);


// --- IDs file Drive ---
const LOGO_FILE_ID = "1Rwo4pJt222dLTXN9W6knN3A5LwJ5TDIa";
const WATERMARK_FILE_ID = "1fNROb-dRtRl2RCCDCxGPozU3oHMSIkHr";

// --- ENV ---
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const GAS_WEBAPP_URL =
    process.env.GAS_WEBAPP_URL ||
    "https://script.google.com/macros/s/AKfycbyYKqYXMlDMG9n_LrpjjNqOtnA6MElh_ds00og0j59-E2UtvGq9YQZVI3lBTUb60Zo-/exec";
const GOOGLE_CREDENTIALS_B64 = process.env.GOOGLE_CREDENTIALS_B64;

if (!SPREADSHEET_ID || !GAS_WEBAPP_URL || !GOOGLE_CREDENTIALS_B64) {
    console.error(
        "❌ Thiếu biến môi trường: SPREADSHEET_ID / GAS_WEBAPP_URL / GOOGLE_CREDENTIALS_B64"
    );
    process.exit(1);
}

// --- Giải mã Service Account JSON ---
const credentials = JSON.parse(
    Buffer.from(GOOGLE_CREDENTIALS_B64, "base64").toString("utf-8")
);
credentials.private_key = credentials.private_key.replace(/\\n/g, "\n").trim();

// --- Google Auth ---
const scopes = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
];
const auth = new google.auth.JWT(
    credentials.client_email,
    null,
    credentials.private_key,
    scopes
);
const sheets = google.sheets({ version: "v4", auth });
const drive = google.drive({ version: "v3", auth });

// --- Express ---
const app = express();
const PORT = process.env.PORT || 3000;

app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));



async function loadDriveImageBase64(fileId) {
    try {
        const meta = await drive.files.get({ fileId, fields: "mimeType" });
        const bin = await drive.files.get(
            { fileId, alt: "media" },
            { responseType: "arraybuffer" }
        );
        const buffer = Buffer.from(bin.data, "binary");
        return `data:${meta.data.mimeType};base64,${buffer.toString("base64")}`;
    } catch (e) {
        console.error(`⚠️ Không tải được file Drive ${fileId}:`, e.message);
        return "";
    }
}

// --- Routes ---
app.get("/", (_req, res) => res.send("🚀 Server chạy ổn! /bbgn để xuất BBGN."));

app.get("/bbgn", async (req, res) => {
    try {
        console.log("▶️ Bắt đầu xuất BBGN ...");

        // --- Lấy mã đơn hàng ---
        const bbgnRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: "file_BBGN_ct!B:B",
        });
        const colB = bbgnRes.data.values ? bbgnRes.data.values.flat() : [];
        const lastRowWithData = colB.length;
        const maDonHang = colB[lastRowWithData - 1];
        if (!maDonHang)
            return res.send(
                "⚠️ Không tìm thấy dữ liệu ở cột B sheet file_BBGN_ct."
            );

        console.log(`✔️ Mã đơn hàng: ${maDonHang} (dòng ${lastRowWithData})`);

        // --- Lấy đơn hàng ---
        const donHangRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: "Don_hang!A1:BJ50000",
        });
        const rows = donHangRes.data.values || [];
        const data = rows.slice(1);
        const donHang =
            data.find((r) => r[5] === maDonHang) ||
            data.find((r) => r[6] === maDonHang);
        if (!donHang)
            return res.send("❌ Không tìm thấy đơn hàng với mã: " + maDonHang);

        // --- Chi tiết sản phẩm ---
        const ctRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: "Don_hang_PVC_ct!A1:AC50000",
        });
        const ctRows = (ctRes.data.values || []).slice(1);
        const products = ctRows
            .filter((r) => r[1] === maDonHang)
            .map((r, i) => ({
                stt: i + 1,
                tenSanPham: r[9],
                soLuong: r[22],
                donVi: r[23],
                tongSoLuong: r[22],
                ghiChu: r[29] || "",
            }));

        console.log(`✔️ Tìm thấy ${products.length} sản phẩm.`);

        // --- Logo & Watermark ---
        const logoBase64 = await loadDriveImageBase64(LOGO_FILE_ID);
        const watermarkBase64 = await loadDriveImageBase64(WATERMARK_FILE_ID);

        // Gọi AppScript để tạo PDF
        const payload = { orderCode: maDonHang };
        const gasResp = await fetch(GAS_WEBAPP_URL, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload),
        });

        // Xử lý response từ AppScript
        const gasText = await gasResp.text();
        let result = {};
        try {
            result = JSON.parse(gasText);
        } catch (e) {
            console.error("❌ Không parse được JSON từ AppScript:", gasText);
            throw new Error("Lỗi từ AppScript: " + gasText);
        }

        if (!result.ok) {
            throw new Error(result.error || "Lỗi từ AppScript");
        }

        console.log("✔️ PDF đã được tạo thành công");

        // Render trang in
        res.render("bbgn", {
            donHang,
            products,
            logoBase64,
            watermarkBase64,
            autoPrint: true,
            maDonHang,
            pathToFile: result.pathToFile || ""
        });
    } catch (err) {
        console.error("❌ Lỗi khi xuất BBGN:", err.stack || err.message);
        res.status(500).send("Lỗi server: " + (err.message || err));
    }
});

// --- Debug ---
app.get("/debug", (_req, res) => {
    res.json({ spreadsheetId: SPREADSHEET_ID, clientEmail: credentials.client_email, gasWebappUrl: GAS_WEBAPP_URL });
});

// --- Start server ---
app.listen(PORT, () => console.log(`✅ Server is running on port ${PORT}`));
