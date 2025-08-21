import dotenv from "dotenv";
import express from "express";
import { google } from "googleapis";
import path from "path";
import pdf from "html-pdf";
import { fileURLToPath } from "url";
import { dirname } from "path";
import ejs from "ejs";

dotenv.config();

// Tạo __dirname trong ESM
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// Load credentials từ biến môi trường
const credentialsBase64 = process.env.GOOGLE_CREDENTIALS_B64;
if (!credentialsBase64) {
    console.error("GOOGLE_CREDENTIALS_B64 environment variable is missing!");
    process.exit(1);
}

const credentials = JSON.parse(Buffer.from(credentialsBase64, "base64").toString("utf-8"));
credentials.private_key = credentials.private_key.replace(/\\n/g, "\n").trim();

// Google Auth
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
const drive = google.drive({ version: "v3", auth });

const app = express();
const PORT = process.env.PORT || 3000;

// EJS view engine
app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));

// Spreadsheet ID
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
if (!SPREADSHEET_ID) {
    console.error("SPREADSHEET_ID environment variable is missing!");
    process.exit(1);
}

const LOGO_FILE_ID = process.env.LOGO_FILE_ID;
const WATERMARK_FILE_ID = process.env.WATERMARK_FILE_ID;
const FOLDER_ID = process.env.FOLDER_ID;

app.get("/", (req, res) => {
    res.send("🚀 Google Sheets API server is running!");
});

app.get("/bbgn", async (req, res) => {
    try {
        // Kiểm tra quyền truy cập
        await auth.getAccessToken();

        // 1. Lấy mã đơn hàng
        const bbgnRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: "file_BBGN_ct!B:B",
        });

        if (!bbgnRes.data.values || bbgnRes.data.values.length === 0) {
            return res.render("error", { message: "Không tìm thấy dữ liệu trong sheet file_BBGN_ct." });
        }

        const colB = bbgnRes.data.values.flat();
        const lastRowWithData = colB.length;
        const maDonHang = colB[lastRowWithData - 1];

        if (!maDonHang) {
            return res.render("error", { message: "Không tìm thấy mã đơn hàng." });
        }

        // 2. Lấy dữ liệu đơn hàng
        const donHangRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: "Don_hang!A:G", // Tối ưu range
        });

        if (!donHangRes.data.values || donHangRes.data.values.length <= 1) {
            return res.render("error", { message: "Không tìm thấy dữ liệu đơn hàng." });
        }

        const rows = donHangRes.data.values;
        const data = rows.slice(1);
        const donHang = data.find((row) => row[6] === maDonHang);

        if (!donHang) {
            return res.render("error", { message: `Không tìm thấy đơn hàng với mã: ${maDonHang}` });
        }

        // 3. Lấy chi tiết sản phẩm
        const ctRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: "Don_hang_PVC_ct!A:AD", // Tối ưu range
        });

        const ctRows = ctRes.data.values.slice(1);
        const products = ctRows
            .filter((row) => row[1] === maDonHang)
            .map((row, index) => ({
                stt: index + 1,
                tenSanPham: row[9],
                soLuong: row[22],
                donVi: row[23],
                tongSoLuong: row[22],
                ghiChu: row[29] || "",
            }));

        // 4. Lấy logo
        let logoBase64 = "";
        try {
            const fileMeta = await drive.files.get({
                fileId: LOGO_FILE_ID,
                fields: "mimeType",
            });
            const resFile = await drive.files.get(
                { fileId: LOGO_FILE_ID, alt: "media" },
                { responseType: "arraybuffer" }
            );
            const buffer = Buffer.from(resFile.data, "binary");
            logoBase64 = `data:${fileMeta.data.mimeType};base64,${buffer.toString("base64")}`;
        } catch (err) {
            console.error("⚠️ Không lấy được logo:", err.message);
        }

        // 5. Lấy watermark
        let watermarkBase64 = "";
        try {
            const fileMeta = await drive.files.get({
                fileId: WATERMARK_FILE_ID,
                fields: "mimeType",
            });
            const resFile = await drive.files.get(
                { fileId: WATERMARK_FILE_ID, alt: "media" },
                { responseType: "arraybuffer" }
            );
            const buffer = Buffer.from(resFile.data, "binary");
            watermarkBase64 = `data:${fileMeta.data.mimeType};base64,${buffer.toString("base64")}`;
        } catch (err) {
            console.error("⚠️ Không lấy được watermark:", err.message);
        }

        // 6. Render HTML từ bbgn.ejs
        const htmlContent = await new Promise((resolve, reject) => {
            app.render("bbgn", {
                donHang,
                products,
                logoBase64,
                watermarkBase64,
                autoPrint: false,
                maDonHang
            }, (err, html) => {
                if (err) reject(err);
                else resolve(html);
            });
        });

        // 7. Dùng html-pdf export ra buffer PDF
        function exportBBGN(htmlContent) {
            return new Promise((resolve, reject) => {
                pdf.create(htmlContent, {
                    format: "A4",
                    border: "10mm",
                    type: "pdf"
                }).toBuffer((err, buffer) => {
                    if (err) reject(err);
                    else resolve(buffer);
                });
            });
        }

        const pdfBuffer = await exportBBGN(htmlContent);

        // 8. Upload PDF lên Google Drive
        const today = new Date();
        const dd = String(today.getDate()).padStart(2, "0");
        const mm = String(today.getMonth() + 1).padStart(2, "0");
        const yyyy = today.getFullYear();
        const hh = String(today.getHours()).padStart(2, "0");
        const mi = String(today.getMinutes()).padStart(2, "0");
        const ss = String(today.getSeconds()).padStart(2, "0");

        const fileName = `BBGN - ${maDonHang} - ${dd}${mm}${yyyy} - ${hh}-${mi}-${ss}.pdf`;
        const fileMeta = { name: fileName, parents: [FOLDER_ID] };
        const media = { mimeType: "application/pdf", body: Buffer.from(pdfBuffer) };

        const pdfFile = await drive.files.create({
            requestBody: fileMeta,
            media,
            fields: "id, name",
        });

        // 9. Lấy tên folder để ghi lại đường dẫn
        const folderMeta = await drive.files.get({
            fileId: FOLDER_ID,
            fields: "name",
        });
        const pathToFile = `${folderMeta.data.name}/${pdfFile.data.name}`;

        // 10. Ghi link file PDF vào Google Sheets
        await sheets.spreadsheets.values.update({
            spreadsheetId: SPREADSHEET_ID,
            range: `file_BBGN_ct!D${lastRowWithData}`,
            valueInputOption: "RAW",
            requestBody: { values: [[pathToFile]] },
        });

        // 11. Render lại bbgn.ejs cho client
        res.render("bbgn", {
            donHang,
            products,
            logoBase64,
            watermarkBase64,
            autoPrint: true,
            maDonHang,
        });
    } catch (err) {
        console.error("Lỗi trong endpoint /bbgn:", err.message);
        res.status(500).render("error", { message: "Đã xảy ra lỗi khi xử lý yêu cầu. Vui lòng thử lại sau." });
    }
});

// Debug endpoint
app.get("/debug", (req, res) => {
    if (process.env.NODE_ENV !== "development") {
        return res.status(403).send("Debug endpoint is disabled in production");
    }
    res.json({
        spreadsheetId: SPREADSHEET_ID,
        scopes: scopes,
    });
});

// Start server
app.listen(PORT, () => {
    console.log(`✅ Server is running on port ${PORT}`);
});