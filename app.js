import express from "express";
import { google } from "googleapis";
import bodyParser from "body-parser";
import fs from "fs";
import path from "path";
import pdf from "html-pdf"; // thay puppeteer bằng html-pdf
import { createRequire } from "module";
const app = express();
app.use(bodyParser.json());
app.set("view engine", "ejs");
const require = createRequire(import.meta.url);
const pdf = require("html-pdf");
// Google API setup
const auth = new google.auth.GoogleAuth({
    credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
    scopes: ["https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/spreadsheets"],
});

const drive = google.drive({ version: "v3", auth });
const sheets = google.sheets({ version: "v4", auth });

const SPREADSHEET_ID = "ID_GOOGLE_SHEET"; // thay ID của bạn

// ✅ Endpoint xuất Biên bản giao nhận
app.get("/bbgn", async (req, res) => {
    try {
        console.log("▶️ Bắt đầu xuất BBGN...");

        // 🔹 1. Lấy dòng cuối cùng trong cột B sheet file_BBGN_ct
        const bbgnRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: "file_BBGN_ct!B:B",
        });

        const bbgnRows = bbgnRes.data.values || [];
        const lastRowIndex = bbgnRows.length; // số dòng cuối
        const maDonHang = bbgnRows[lastRowIndex - 1][0];

        if (!maDonHang) {
            return res.send("⚠️ Không tìm thấy mã đơn hàng trong file_BBGN_ct!");
        }

        console.log("✅ Mã đơn hàng cuối:", maDonHang);

        // 🔹 2. Lấy dữ liệu đơn hàng từ sheet Don_hang
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

        // 🔹 3. Lấy chi tiết sản phẩm
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

        console.log(`✅ Tìm thấy ${products.length} sản phẩm`);

        // 🔹 4. Render HTML từ EJS
        res.render("bbgn", { donHang, products, autoPrint: false }, async (err, html) => {
            if (err) {
                console.error("❌ Lỗi render EJS:", err);
                return res.status(500).send("Lỗi render");
            }

            // 🔹 5. Xuất PDF từ HTML
            const now = new Date();
            const dateStr = now.toLocaleDateString("vi-VN").replace(/\//g, "");
            const timeStr = now.toTimeString().split(" ")[0].replace(/:/g, "-");
            const fileName = `BBGN - ${maDonHang} - ${dateStr} - ${timeStr}.pdf`;

            const pdfPath = path.join("/tmp", fileName); // Render free chỉ ghi tạm ở /tmp

            pdf.create(html, { format: "A4" }).toFile(pdfPath, async (err, pdfRes) => {
                if (err) {
                    console.error("❌ Lỗi tạo PDF:", err);
                    return res.status(500).send("Lỗi khi xuất PDF");
                }

                console.log("✅ PDF đã tạo:", pdfRes.filename);

                // 🔹 6. Upload PDF lên Google Drive
                const folderId = "1CL3JuFprNj1a406XWXTtbQMZmyKxhczW";
                const fileMetadata = {
                    name: fileName,
                    parents: [folderId],
                };

                const media = {
                    mimeType: "application/pdf",
                    body: fs.createReadStream(pdfPath),
                };

                const uploadedFile = await drive.files.create({
                    resource: fileMetadata,
                    media,
                    fields: "id, name",
                });

                console.log("✅ File đã upload:", uploadedFile.data);

                // 🔹 7. Ghi đường dẫn vào cột D cùng dòng đó trong sheet file_BBGN_ct
                const folderMeta = await drive.files.get({
                    fileId: folderId,
                    fields: "name",
                });

                const folderName = folderMeta.data.name;
                const pathToFile = `${folderName}/${fileName}`;

                await sheets.spreadsheets.values.update({
                    spreadsheetId: SPREADSHEET_ID,
                    range: `file_BBGN_ct!D${lastRowIndex}`,
                    valueInputOption: "USER_ENTERED",
                    requestBody: {
                        values: [[pathToFile]],
                    },
                });

                console.log("✅ Đã ghi đường dẫn vào sheet:", pathToFile);

                res.send(`✅ Đã tạo và lưu BBGN thành công! File: ${pathToFile}`);
            });
        });
    } catch (err) {
        console.error("❌ Lỗi xuất BBGN:", err);
        res.status(500).send("Lỗi hệ thống khi xuất BBGN");
    }
});
