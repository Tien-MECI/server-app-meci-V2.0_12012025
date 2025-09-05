import dotenv from "dotenv";
import express from "express";
import { google } from "googleapis";
import path from "path";
import { fileURLToPath } from "url";
import { dirname } from "path";
import ejs from "ejs";
import fetch from "node-fetch";
import { promisify } from "util";
const renderFileAsync = promisify(ejs.renderFile);


dotenv.config();

// --- __dirname trong ESM ---
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);


// --- IDs file Drive ---
const LOGO_FILE_ID = "1Rwo4pJt222dLTXN9W6knN3A5LwJ5TDIa";
const WATERMARK_FILE_ID = "1fNROb-dRtRl2RCCDCxGPozU3oHMSIkHr";

// --- ENV ---
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const SPREADSHEET_HC_ID = process.env.SPREADSHEET_HC_ID;
const GAS_WEBAPP_URL = process.env.GAS_WEBAPP_URL;
const GAS_WEBAPP_URL_BBNT = process.env.GAS_WEBAPP_URL_BBNT;
const GOOGLE_CREDENTIALS_B64 = process.env.GOOGLE_CREDENTIALS_B64;

if (!SPREADSHEET_ID || !SPREADSHEET_HC_ID ||!GAS_WEBAPP_URL || !GAS_WEBAPP_URL_BBNT || !GOOGLE_CREDENTIALS_B64) {
    console.error(
        "❌ Thiếu biến môi trường: SPREADSHEET_ID / SPREADSHEET_HC_ID / GAS_WEBAPP_URL / GAS_WEBAPP_URL_BBNT / GOOGLE_CREDENTIALS_B64"
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
            return res.send("⚠️ Không tìm thấy dữ liệu ở cột B sheet file_BBGN_ct.");

        console.log(`✔️ Mã đơn hàng: ${maDonHang} (dòng ${lastRowWithData})`);

        // --- Lấy đơn hàng ---
        const donHangRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: "Don_hang!A1:BJ",
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
            range: "Don_hang_PVC_ct!A1:AC",
        });
        const ctRows = (ctRes.data.values || []).slice(1);
        const products = ctRows
            .filter((r) => r[1] === maDonHang)
            .map((r, i) => ({
                stt: i + 1,
                tenSanPham: r[9],
                soLuong: r[23],
                donVi: r[22],
                tongSoLuong: r[21],
                ghiChu: "",
            }));

        console.log(`✔️ Tìm thấy ${products.length} sản phẩm.`);

        // --- Logo & Watermark ---
        const logoBase64 = await loadDriveImageBase64(LOGO_FILE_ID);
        const watermarkBase64 = await loadDriveImageBase64(WATERMARK_FILE_ID);

        // --- Render ngay cho client ---
        res.render("bbgn", {
            donHang,
            products,
            logoBase64,
            watermarkBase64,
            autoPrint: true,
            maDonHang,
            pathToFile: ""
        });

        // --- Sau khi render xong thì gọi AppScript ngầm ---
        (async () => {
            try {
                const renderedHtml = await renderFileAsync(
                    path.join(__dirname, "views", "bbgn.ejs"),
                    {
                        donHang,
                        products,
                        logoBase64,
                        watermarkBase64,
                        autoPrint: false,
                        maDonHang,
                        pathToFile: ""
                    }
                );

                const resp = await fetch(GAS_WEBAPP_URL, {
                    method: "POST",
                    headers: { "Content-Type": "application/x-www-form-urlencoded" },
                    body: new URLSearchParams({
                        orderCode: maDonHang,
                        html: renderedHtml
                    })
                });

                const data = await resp.json();
                console.log("✔️ AppScript trả về:", data);

                const pathToFile = data.pathToFile || `BBGN/${data.fileName}`;
                await sheets.spreadsheets.values.update({
                    spreadsheetId: SPREADSHEET_ID,
                    range: `file_BBGN_ct!D${lastRowWithData}`,
                    valueInputOption: "RAW",
                    requestBody: { values: [[pathToFile]] },
                });
                console.log("✔️ Đã ghi đường dẫn:", pathToFile);

            } catch (err) {
                console.error("❌ Lỗi gọi AppScript:", err);
            }
        })();

    } catch (err) {
        console.error("❌ Lỗi khi xuất BBGN:", err.stack || err.message);
        res.status(500).send("Lỗi server: " + (err.message || err));
    }
});


app.get("/bbnt", async (req, res) => {
    try {
        console.log("▶️ Bắt đầu xuất BBNT ...");

        // --- Lấy mã đơn hàng ---
        const bbntRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: "File_BBNT_ct!B:B",
        });
        const colB = bbntRes.data.values ? bbntRes.data.values.flat() : [];
        const lastRowWithData = colB.length;
        const maDonHang = colB[lastRowWithData - 1];
        if (!maDonHang)
            return res.send("⚠️ Không tìm thấy dữ liệu ở cột B sheet File_BBNT_ct.");

        console.log(`✔️ Mã đơn hàng: ${maDonHang} (dòng ${lastRowWithData})`);

        // --- Lấy đơn hàng ---
        const donHangRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: "Don_hang!A1:BJ",
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
            range: "Don_hang_PVC_ct!A1:AC",
        });
        const ctRows = (ctRes.data.values || []).slice(1);
        const products = ctRows
            .filter((r) => r[1] === maDonHang)
            .map((r, i) => ({
                stt: i + 1,
                tenSanPham: r[9],
                soLuong: r[23],
                donVi: r[22],
                tongSoLuong: r[21],
                ghiChu: "",
            }));

        console.log(`✔️ Tìm thấy ${products.length} sản phẩm.`);

        // --- Logo & Watermark ---
        const logoBase64 = await loadDriveImageBase64(LOGO_FILE_ID);
        const watermarkBase64 = await loadDriveImageBase64(WATERMARK_FILE_ID);

        // --- Render ngay cho client ---
        res.render("bbnt", {
            donHang,
            products,
            logoBase64,
            watermarkBase64,
            autoPrint: true,
            maDonHang,
            pathToFile: ""
        });

        // --- Sau khi render xong thì gọi AppScript ngầm ---
        (async () => {
            try {
                const renderedHtml = await renderFileAsync(
                    path.join(__dirname, "views", "bbnt.ejs"),
                    {
                        donHang,
                        products,
                        logoBase64,
                        watermarkBase64,
                        autoPrint: false,
                        maDonHang,
                        pathToFile: ""
                    }
                );

                const resp = await fetch(GAS_WEBAPP_URL_BBNT, {
                    method: "POST",
                    headers: { "Content-Type": "application/x-www-form-urlencoded" },
                    body: new URLSearchParams({
                        orderCode: maDonHang,
                        html: renderedHtml
                    })
                });

                const data = await resp.json();
                console.log("✔️ AppScript trả về:", data);

                const pathToFile = data.pathToFile || `BBNT/${data.fileName}`;
                await sheets.spreadsheets.values.update({
                    spreadsheetId: SPREADSHEET_ID,
                    range: `File_BBNT_ct!D${lastRowWithData}`,
                    valueInputOption: "RAW",
                    requestBody: { values: [[pathToFile]] },
                });
                console.log("✔️ Đã ghi đường dẫn:", pathToFile);

            } catch (err) {
                console.error("❌ Lỗi gọi AppScript:", err);
            }
        })();

    } catch (err) {
        console.error("❌ Lỗi chi tiết:", err);
        res.status(500).send("Lỗi server: " + err.message);
    }
});


app.get("/ggh", async (req, res) => {
    try {
        console.log("▶️ Bắt đầu xuất GGH ...");

        // --- Lấy mã đơn hàng ---
        const gghRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: "File_GGH_ct!B:B",
        });
        const colB = gghRes.data.values ? gghRes.data.values.flat() : [];
        const lastRowWithData = colB.length;
        const maDonHang = colB[lastRowWithData - 1];
        if (!maDonHang)
            return res.send("⚠️ Không tìm thấy dữ liệu ở cột B sheet File_GGH_ct.");

        console.log(`✔️ Mã đơn hàng: ${maDonHang} (dòng ${lastRowWithData})`);

        // --- Lấy đơn hàng ---
        const donHangRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: "Don_hang!A1:BJ",
        });
        const rows = donHangRes.data.values || [];
        const data = rows.slice(1);
        const donHang =
            data.find((r) => r[5] === maDonHang) ||
            data.find((r) => r[6] === maDonHang);
        if (!donHang)
            return res.send("❌ Không tìm thấy đơn hàng với mã: " + maDonHang);

        // --- Logo ---
        const logoBase64 = await loadDriveImageBase64(LOGO_FILE_ID);

        // --- Render ngay cho client ---
        res.render("ggh", {
            donHang,
            logoBase64,
            autoPrint: false,
            maDonHang,
            pathToFile: ""
        });

        // --- Sau khi render xong thì gọi AppScript ngầm ---
        (async () => {
            try {
                const renderedHtml = await renderFileAsync(
                    path.join(__dirname, "views", "ggh.ejs"),
                    {
                        donHang,
                        logoBase64,
                        autoPrint: false,
                        maDonHang,
                        pathToFile: ""
                    }
                );

                // Gọi GAS webapp tương ứng (cần thêm biến môi trường GAS_WEBAPP_URL_GGH)
                const GAS_WEBAPP_URL_GGH = process.env.GAS_WEBAPP_URL_GGH;
                if (GAS_WEBAPP_URL_GGH) {
                    const resp = await fetch(GAS_WEBAPP_URL_GGH, {
                        method: "POST",
                        headers: { "Content-Type": "application/x-www-form-urlencoded" },
                        body: new URLSearchParams({
                            orderCode: maDonHang,
                            html: renderedHtml
                        })
                    });

                    const data = await resp.json();
                    console.log("✔️ AppScript trả về:", data);

                    const pathToFile = data.pathToFile || `GGH/${data.fileName}`;
                    await sheets.spreadsheets.values.update({
                        spreadsheetId: SPREADSHEET_ID,
                        range: `File_GGH_ct!D${lastRowWithData}`,
                        valueInputOption: "RAW",
                        requestBody: { values: [[pathToFile]] },
                    });
                    console.log("✔️ Đã ghi đường dẫn:", pathToFile);
                } else {
                    console.log("⚠️ Chưa cấu hình GAS_WEBAPP_URL_GGH");
                }

            } catch (err) {
                console.error("❌ Lỗi gọi AppScript:", err);
            }
        })();

    } catch (err) {
        console.error("❌ Lỗi khi xuất GGH:", err.stack || err.message);
        res.status(500).send("Lỗi server: " + (err.message || err));
    }
});


app.get("/lenhpvc", async (req, res) => {
    try {
        console.log("▶️ Bắt đầu xuất Lệnh PVC ...");

        // --- Lấy mã đơn hàng ---
        const lenhRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: "File_lenh_ct!B:B",
        });
        const colB = lenhRes.data.values ? lenhRes.data.values.flat() : [];
        const lastRowWithData = colB.length;
        const maDonHang = colB[lastRowWithData - 1];
        if (!maDonHang)
            return res.send("⚠️ Không tìm thấy dữ liệu ở cột B sheet File_lenh_ct.");

        console.log(`✔️ Mã đơn hàng: ${maDonHang} (dòng ${lastRowWithData})`);

        // --- Lấy đơn hàng ---
        const donHangRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: "Don_hang!A1:BJ",
        });
        const rows = donHangRes.data.values || [];
        const data = rows.slice(1);
        const donHang =
            data.find((r) => r[5] === maDonHang) ||
            data.find((r) => r[6] === maDonHang);
        if (!donHang)
            return res.send("❌ Không tìm thấy đơn hàng với mã: " + maDonHang);

        // --- Lấy chi tiết sản phẩm PVC ---
        const ctRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: "Don_hang_PVC_ct!A1:AC",
        });
        const ctRows = (ctRes.data.values || []).slice(1);

        // Lọc và map dữ liệu theo cấu trúc của lệnh sản xuất
        const products = ctRows
            .filter((r) => r[1] === maDonHang)
            .map((r, i) => ({
                maDonHangChiTiet: r[2],
                tenThuongMai: r[9],
                dai: r[16],
                rong: r[17],
                cao: r[18],
                slSoi: r[19],
                soLuong: r[21],
                donViTinh: r[22],
                tongSoLuong: r[20],
                tongSLSoi: r[23],
                ghiChuSanXuat: r[28]
            }));

        console.log(`✔️ Tìm thấy ${products.length} sản phẩm.`);

        // --- Logo & Watermark ---
        const logoBase64 = await loadDriveImageBase64(LOGO_FILE_ID);
        const watermarkBase64 = await loadDriveImageBase64(WATERMARK_FILE_ID);

        // --- Xác định loại lệnh từ cột S (index 36) ---
        const lenhValue = donHang[36] || '';

        // --- Render ngay cho client ---
        res.render("lenhpvc", {
            donHang,
            products,
            logoBase64,
            watermarkBase64,
            autoPrint: false,
            maDonHang,
            lenhValue,
            pathToFile: ""
        });

        // --- Sau khi render xong thì gọi AppScript ngầm ---
        (async () => {
            try {
                const renderedHtml = await renderFileAsync(
                    path.join(__dirname, "views", "lenhpvc.ejs"),
                    {
                        donHang,
                        products,
                        logoBase64,
                        watermarkBase64,
                        autoPrint: false,
                        maDonHang,
                        lenhValue,
                        pathToFile: ""
                    }
                );

                // Gọi GAS webapp tương ứng (cần thêm biến môi trường GAS_WEBAPP_URL_LENHPVC)
                const GAS_WEBAPP_URL_LENHPVC = process.env.GAS_WEBAPP_URL_LENHPVC;
                if (GAS_WEBAPP_URL_LENHPVC) {
                    const resp = await fetch(GAS_WEBAPP_URL_LENHPVC, {
                        method: "POST",
                        headers: { "Content-Type": "application/x-www-form-urlencoded" },
                        body: new URLSearchParams({
                            orderCode: maDonHang,
                            html: renderedHtml
                        })
                    });

                    const data = await resp.json();
                    console.log("✔️ AppScript trả về:", data);

                    const pathToFile = data.pathToFile || `LENH_PVC/${data.fileName}`;
                    await sheets.spreadsheets.values.update({
                        spreadsheetId: SPREADSHEET_ID,
                        range: `File_lenh_ct!D${lastRowWithData}`,
                        valueInputOption: "RAW",
                        requestBody: { values: [[pathToFile]] },
                    });
                    console.log("✔️ Đã ghi đường dẫn:", pathToFile);
                } else {
                    console.log("⚠️ Chưa cấu hình GAS_WEBAPP_URL_LENHPVC");
                }

            } catch (err) {
                console.error("❌ Lỗi gọi AppScript:", err);
            }
        })();

    } catch (err) {
        console.error("❌ Lỗi khi xuất Lệnh PVC:", err.stack || err.message);
        res.status(500).send("Lỗi server: " + (err.message || err));
    }
});

app.get("/baogiapvc", async (req, res) => {
    try {
        console.log("▶️ Bắt đầu xuất Báo Giá PVC ...");

        // --- Lấy mã đơn hàng ---
        const baoGiaRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: "File_bao_gia_ct!B:B",
        });
        const colB = baoGiaRes.data.values ? baoGiaRes.data.values.flat() : [];
        const lastRowWithData = colB.length;
        const maDonHang = colB[lastRowWithData - 1];
        if (!maDonHang)
            return res.send("⚠️ Không tìm thấy dữ liệu ở cột B sheet File_bao_gia_ct.");

        console.log(`✔️ Mã đơn hàng: ${maDonHang} (dòng ${lastRowWithData})`);

        // --- Lấy đơn hàng ---
        const donHangRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: "Don_hang!A1:BJ",
        });
        const rows = donHangRes.data.values || [];
        const data = rows.slice(1);
        const donHang =
            data.find((r) => r[5] === maDonHang) ||
            data.find((r) => r[6] === maDonHang);
        if (!donHang)
            return res.send("❌ Không tìm thấy đơn hàng với mã: " + maDonHang);

        // --- Lấy chi tiết sản phẩm PVC ---
        const ctRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: "Don_hang_PVC_ct!A1:AC",
        });
        const ctRows = (ctRes.data.values || []).slice(1);

        // Lọc và map dữ liệu theo cấu trúc của báo giá
        const products = ctRows
            .filter((r) => r[1] === maDonHang)
            .map((r) => ({
                maDonHangChiTiet: r[2],
                tenHangHoa: r[9],
                quyCach: r[10],
                dai: r[16],
                rong: r[17],
                cao: r[18],
                soLuong: r[21],
                donViTinh: r[22],
                tongSoLuong: r[20],
                donGia: r[25],
                vat: r[26] ? parseFloat(r[26]) : null,
                thanhTien: r[27]
            }));

        console.log(`✔️ Tìm thấy ${products.length} sản phẩm.`);

        // --- Tính tổng các giá trị ---
        let tongTien = 0;
        let chietKhau = parseFloat(donHang[40]) || 0;
        let tamUng = parseFloat(donHang[41]) || 0;

        products.forEach(product => {
            tongTien += parseFloat(product.thanhTien) || 0;
        });

        let tongThanhTien = tongTien - chietKhau - tamUng;

        // --- Logo & Watermark ---
        const logoBase64 = await loadDriveImageBase64(LOGO_FILE_ID);
        const watermarkBase64 = await loadDriveImageBase64(WATERMARK_FILE_ID);

        // --- Render ngay cho client ---
        res.render("baogiapvc", {
            donHang,
            products,
            logoBase64,
            watermarkBase64,
            autoPrint: false,
            maDonHang,
            tongTien,
            chietKhau,
            tamUng,
            tongThanhTien,
            numberToWords,
            pathToFile: ""
            
        });

        // --- Sau khi render xong thì gọi AppScript ngầm ---
        (async () => {
            try {
                const renderedHtml = await renderFileAsync(
                    path.join(__dirname, "views", "baogiapvc.ejs"),
                    {
                        donHang,
                        products,
                        logoBase64,
                        watermarkBase64,
                        autoPrint: false,
                        maDonHang,
                        tongTien,
                        chietKhau,
                        tamUng,
                        tongThanhTien,
                        numberToWords,
                        pathToFile: ""
                    }
                );

                // Gọi GAS webapp tương ứng (cần thêm biến môi trường GAS_WEBAPP_URL_BAOGIA)
                const GAS_WEBAPP_URL_BAOGIAPVC = process.env.GAS_WEBAPP_URL_BAOGIAPVC;
                if (GAS_WEBAPP_URL_BAOGIAPVC) {
                    const resp = await fetch(GAS_WEBAPP_URL_BAOGIAPVC, {
                        method: "POST",
                        headers: { "Content-Type": "application/x-www-form-urlencoded" },
                        body: new URLSearchParams({
                            orderCode: maDonHang,
                            html: renderedHtml
                        })
                    });

                    const data = await resp.json();
                    console.log("✔️ AppScript trả về:", data);

                    const pathToFile = data.pathToFile || `BAO_GIA_PVC/${data.fileName}`;
                    await sheets.spreadsheets.values.update({
                        spreadsheetId: SPREADSHEET_ID,
                        range: `File_bao_gia_ct!D${lastRowWithData}`,
                        valueInputOption: "RAW",
                        requestBody: { values: [[pathToFile]] },
                    });
                    console.log("✔️ Đã ghi đường dẫn:", pathToFile);
                } else {
                    console.log("⚠️ Chưa cấu hình GAS_WEBAPP_URL_BAOGIA");
                }

            } catch (err) {
                console.error("❌ Lỗi gọi AppScript:", err);
            }
        })();

    } catch (err) {
        console.error("❌ Lỗi khi xuất Báo Giá PVC:", err.stack || err.message);
        res.status(500).send("Lỗi server: " + (err.message || err));
    }
});

app.get("/baogiank", async (req, res) => {
    try {
        console.log("▶️ Bắt đầu xuất Báo Giá Nhôm Kính ...");

        // --- Lấy mã đơn hàng ---
        const baoGiaRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: "File_bao_gia_ct!B:B",
        });
        const colB = baoGiaRes.data.values ? baoGiaRes.data.values.flat() : [];
        const lastRowWithData = colB.length;
        const maDonHang = colB[lastRowWithData - 1];
        if (!maDonHang)
            return res.send("⚠️ Không tìm thấy dữ liệu ở cột B sheet File_bao_gia_ct.");

        console.log(`✔️ Mã đơn hàng: ${maDonHang} (dòng ${lastRowWithData})`);

        // --- Lấy đơn hàng ---
        const donHangRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: "Don_hang!A1:BW", // Mở rộng đến cột BW
        });
        const rows = donHangRes.data.values || [];
        const data = rows.slice(1);
        const donHang =
            data.find((r) => r[5] === maDonHang) ||
            data.find((r) => r[6] === maDonHang);
        if (!donHang)
            return res.send("❌ Không tìm thấy đơn hàng với mã: " + maDonHang);

        // --- Lấy chi tiết sản phẩm Nhôm Kính ---
        const ctRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: "Don_hang_nk_ct!A1:U", // Mở rộng đến cột U
        });
        const ctRows = (ctRes.data.values || []).slice(1);
        
        // Lọc và map dữ liệu theo cấu trúc của báo giá nhôm kính
        const products = ctRows
            .filter((r) => r[1] === maDonHang)
            .map((r) => ({
                kyHieu: r[5],
                tenHangHoa: r[8],
                dai: r[9],
                rong: r[10],
                cao: r[11],
                dienTich: r[12],
                soLuong: r[14],
                donViTinh: r[13],
                donGia: r[17],
                giaPK: r[16],
                thanhTien: r[19]
            }));

        console.log(`✔️ Tìm thấy ${products.length} sản phẩm.`);

        // --- Tính tổng các giá trị ---
        let tongTien = 0;
        let chietKhau = parseFloat(donHang[40]) || 0; // Cột AN
        let tamUng = parseFloat(donHang[41]) || 0; // Cột AO
        
        products.forEach(product => {
            tongTien += parseFloat(product.thanhTien) || 0;
        });

        let tongThanhTien = tongTien - chietKhau - tamUng;

        // Tính tổng diện tích và số lượng
        let tongDienTich = 0;
        let tongSoLuong = 0;
        
        products.forEach(product => {
            const dienTich = parseFloat(product.dienTich) || 0;
            const soLuong = parseFloat(product.soLuong) || 0;
            tongDienTich += dienTich * soLuong;
            tongSoLuong += soLuong;
        });

        tongDienTich = parseFloat(tongDienTich.toFixed(2));

        // --- Logo & Watermark ---
        const logoBase64 = await loadDriveImageBase64(LOGO_FILE_ID);
        const watermarkBase64 = await loadDriveImageBase64('1766zFeBWPEmjTGQGrrtM34QFbV8fHryb'); // Watermark ID từ code GAS

        // --- Render ngay cho client ---
        res.render("baogiank", {
            donHang,
            products,
            logoBase64,
            watermarkBase64,
            autoPrint: true,
            maDonHang,
            tongTien,
            chietKhau,
            tamUng,
            tongThanhTien,
            tongDienTich,
            tongSoLuong,
            numberToWords,
            pathToFile: ""
        });

        // --- Sau khi render xong thì gọi AppScript ngầm ---
        (async () => {
            try {
                const renderedHtml = await renderFileAsync(
                    path.join(__dirname, "views", "baogiank.ejs"),
                    {
                        donHang,
                        products,
                        logoBase64,
                        watermarkBase64,
                        autoPrint: false,
                        maDonHang,
                        tongTien,
                        chietKhau,
                        tamUng,
                        tongThanhTien,
                        tongDienTich,
                        tongSoLuong,
                        numberToWords,
                        pathToFile: ""
                    }
                );

                // Gọi GAS webapp tương ứng (cần thêm biến môi trường GAS_WEBAPP_URL_BAOGIANK)
                const GAS_WEBAPP_URL_BAOGIANK = process.env.GAS_WEBAPP_URL_BAOGIANK;
                if (GAS_WEBAPP_URL_BAOGIANK) {
                    const resp = await fetch(GAS_WEBAPP_URL_BAOGIANK, {
                        method: "POST",
                        headers: { "Content-Type": "application/x-www-form-urlencoded" },
                        body: new URLSearchParams({
                            orderCode: maDonHang,
                            html: renderedHtml
                        })
                    });

                    const data = await resp.json();
                    console.log("✔️ AppScript trả về:", data);

                    const pathToFile = data.pathToFile || `BAO_GIA_NK/${data.fileName}`;
                    await sheets.spreadsheets.values.update({
                        spreadsheetId: SPREADSHEET_ID,
                        range: `File_bao_gia_ct!D${lastRowWithData}`,
                        valueInputOption: "RAW",
                        requestBody: { values: [[pathToFile]] },
                    });
                    console.log("✔️ Đã ghi đường dẫn:", pathToFile);
                } else {
                    console.log("⚠️ Chưa cấu hình GAS_WEBAPP_URL_BAOGIANK");
                }

            } catch (err) {
                console.error("❌ Lỗi gọi AppScript:", err);
            }
        })();

    } catch (err) {
        console.error("❌ Lỗi khi xuất Báo Giá Nhôm Kính:", err.stack || err.message);
        res.status(500).send("Lỗi server: " + (err.message || err));
    }
});

app.get("/lenhnk", async (req, res) => {
    try {
        console.log("▶️ Bắt đầu xuất Lệnh Nhôm Kính ...");

        // --- Lấy mã đơn hàng ---
        const lenhRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: "File_lenh_ct!B:B",
        });
        const colB = lenhRes.data.values ? lenhRes.data.values.flat() : [];
        const lastRowWithData = colB.length;
        const maDonHang = colB[lastRowWithData - 1];
        if (!maDonHang)
            return res.send("⚠️ Không tìm thấy dữ liệu ở cột B sheet File_lenh_ct.");

        console.log(`✔️ Mã đơn hàng: ${maDonHang} (dòng ${lastRowWithData})`);

        // --- Lấy đơn hàng ---
        const donHangRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: "Don_hang!A1:BJ",
        });
        const rows = donHangRes.data.values || [];
        const data = rows.slice(1);
        const donHang =
            data.find((r) => r[5] === maDonHang) ||
            data.find((r) => r[6] === maDonHang);
        if (!donHang)
            return res.send("❌ Không tìm thấy đơn hàng với mã: " + maDonHang);

        // --- Lấy chi tiết sản phẩm Nhôm Kính ---
        const ctRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: "Don_hang_nk_ct!A1:U",
        });
        const ctRows = (ctRes.data.values || []).slice(1);
        
        // Lọc và map dữ liệu theo cấu trúc của lệnh sản xuất nhôm kính
        const products = ctRows
            .filter((r) => r[1] === maDonHang)
            .map((r, i) => ({
                maDonHangChiTiet: r[2],
                tenThuongMai: r[7],
                dai: r[9],
                rong: r[10],
                cao: r[11],
                dienTich: r[12],
                donViTinh: r[13],
                slBo: r[14],
                tongSoLuong: r[15],
                ghiChuSanXuat: r[20]
            }));

        console.log(`✔️ Tìm thấy ${products.length} sản phẩm.`);

        // --- Logo & Watermark ---
        const logoBase64 = await loadDriveImageBase64(LOGO_FILE_ID);
        const watermarkBase64 = await loadDriveImageBase64(WATERMARK_FILE_ID);

        // --- Xác định loại lệnh từ cột S (index 36) ---
        const lenhValue = donHang[36] || '';

        // --- Render ngay cho client ---
        res.render("lenhnk", {
            donHang,
            products,
            logoBase64,
            watermarkBase64,
            autoPrint: true,
            maDonHang,
            lenhValue,
            pathToFile: ""
        });

        // --- Sau khi render xong thì gọi AppScript ngầm ---
        (async () => {
            try {
                const renderedHtml = await renderFileAsync(
                    path.join(__dirname, "views", "lenhnk.ejs"),
                    {
                        donHang,
                        products,
                        logoBase64,
                        watermarkBase64,
                        autoPrint: false,
                        maDonHang,
                        lenhValue,
                        pathToFile: ""
                    }
                );

                // Gọi GAS webapp tương ứng (cần thêm biến môi trường GAS_WEBAPP_URL_LENHNK)
                const GAS_WEBAPP_URL_LENHNK = process.env.GAS_WEBAPP_URL_LENHNK;
                if (GAS_WEBAPP_URL_LENHNK) {
                    const resp = await fetch(GAS_WEBAPP_URL_LENHNK, {
                        method: "POST",
                        headers: { "Content-Type": "application/x-www-form-urlencoded" },
                        body: new URLSearchParams({
                            orderCode: maDonHang,
                            html: renderedHtml
                        })
                    });

                    const data = await resp.json();
                    console.log("✔️ AppScript trả về:", data);

                    const pathToFile = data.pathToFile || `LENH_NK/${data.fileName}`;
                    await sheets.spreadsheets.values.update({
                        spreadsheetId: SPREADSHEET_ID,
                        range: `File_lenh_ct!D${lastRowWithData}`,
                        valueInputOption: "RAW",
                        requestBody: { values: [[pathToFile]] },
                    });
                    console.log("✔️ Đã ghi đường dẫn:", pathToFile);
                } else {
                    console.log("⚠️ Chưa cấu hình GAS_WEBAPP_URL_LENHNK");
                }

            } catch (err) {
                console.error("❌ Lỗi gọi AppScript:", err);
            }
        })();

    } catch (err) {
        console.error("❌ Lỗi khi xuất Lệnh Nhôm Kính:", err.stack || err.message);
        res.status(500).send("Lỗi server: " + (err.message || err));
    }
});


app.get("/bbgnnk", async (req, res) => {
    try {
        console.log("▶️ Bắt đầu xuất BBGN NK ...");

        // --- Lấy mã đơn hàng ---
        const bbgnnkRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: "file_BBGN_ct!B:B",
        });
        const colB = bbgnnkRes.data.values ? bbgnnkRes.data.values.flat() : [];
        const lastRowWithData = colB.length;
        const maDonHang = colB[lastRowWithData - 1];
        if (!maDonHang) {
            return res.send("⚠️ Không tìm thấy dữ liệu ở cột B sheet file_BBGN_ct.");
        }

        console.log(`✔️ Mã đơn hàng: ${maDonHang} (dòng ${lastRowWithData})`);

        // --- Lấy đơn hàng ---
        const donHangRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: "Don_hang!A1:BJ",
        });
        const rows = donHangRes.data.values || [];
        const data = rows.slice(1);
        const donHang =
            data.find((r) => r[5] === maDonHang) ||
            data.find((r) => r[6] === maDonHang);
        if (!donHang) {
            return res.send("❌ Không tìm thấy đơn hàng với mã: " + maDonHang);
        }

        // --- Chi tiết sản phẩm ---
        const ctRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: "Don_hang_nk_ct!A1:AC",
        });
        const ctRows = (ctRes.data.values || []).slice(1);
        const products = ctRows
            .filter((r) => r[1] === maDonHang)
            .map((r, i) => ({
                stt: i + 1,
                tenSanPham: r[8],
                soLuong: r[14],
                donVi: r[13],
                tongSoLuong: r[15],
                ghiChu: " ",
            }));

        console.log(`✔️ Tìm thấy ${products.length} sản phẩm.`);

        // --- Logo & Watermark ---
        const logoBase64 = await loadDriveImageBase64(LOGO_FILE_ID);
        const watermarkBase64 = await loadDriveImageBase64(WATERMARK_FILE_ID);

        // --- Render ngay cho client ---
        res.render("bbgnnk", {
            donHang,
            products,
            logoBase64,
            watermarkBase64,
            autoPrint: true,
            maDonHang,
            pathToFile: "",
        });

        // --- Sau khi render xong thì gọi AppScript ngầm ---
        (async () => {
            try {
                const renderedHtml = await renderFileAsync(
                    path.join(__dirname, "views", "bbgnnk.ejs"),
                    {
                        donHang,
                        products,
                        logoBase64,
                        watermarkBase64,
                        autoPrint: false,
                        maDonHang,
                        pathToFile: "",
                    }
                );

                const GAS_WEBAPP_URL_BBGNNK = process.env.GAS_WEBAPP_URL_BBGNNK;
                if (GAS_WEBAPP_URL_BBGNNK) {
                    const resp = await fetch(GAS_WEBAPP_URL_BBGNNK, {
                        method: "POST",
                        headers: { "Content-Type": "application/x-www-form-urlencoded" },
                        body: new URLSearchParams({
                            orderCode: maDonHang,
                            html: renderedHtml,
                        }),
                    });

                    const data = await resp.json();
                    console.log("✔️ AppScript trả về:", data);

                    const pathToFile = data.pathToFile || `BBGNNK/${data.fileName}`;
                    await sheets.spreadsheets.values.update({
                        spreadsheetId: SPREADSHEET_ID,
                        range: `file_BBGN_ct!D${lastRowWithData}`,
                        valueInputOption: "RAW",
                        requestBody: { values: [[pathToFile]] },
                    });
                    console.log("✔️ Đã ghi đường dẫn:", pathToFile);
                }
            } catch (err) {
                console.error("❌ Lỗi gọi AppScript:", err);
            }
        })();

    } catch (err) {
        console.error("❌ Lỗi khi xuất BBGN NK:", err.stack || err.message);
        res.status(500).send("Lỗi server: " + (err.message || err));
    }
});

app.get("/bbntnk", async (req, res) => {
  try {
    console.log("▶️ Bắt đầu xuất BBNTNK ...");

    // 1. Lấy mã đơn hàng từ sheet file_BBNT_ct
    const bbntRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "file_BBNT_ct!B:B",
    });
    const colB = bbntRes.data.values ? bbntRes.data.values.flat() : [];
    const lastRowWithData = colB.length;
    const maDonHang = colB[lastRowWithData - 1];
    if (!maDonHang) return res.send("⚠️ Không tìm thấy dữ liệu ở cột B sheet file_BBNT_ct.");

    console.log(`✔️ Mã đơn hàng: ${maDonHang} (dòng ${lastRowWithData})`);

    // 2. Lấy đơn hàng
    const donHangRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Don_hang!A1:BJ",
    });
    const rows = donHangRes.data.values || [];
    const data = rows.slice(1);
    const donHang =
      data.find((r) => r[5] === maDonHang) || data.find((r) => r[6] === maDonHang);
    if (!donHang) return res.send("❌ Không tìm thấy đơn hàng với mã: " + maDonHang);

    // 3. Lấy chi tiết sản phẩm
    const ctRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Don_hang_nk_ct!A1:AC",
    });
    const ctRows = (ctRes.data.values || []).slice(1);
    const products = ctRows
      .filter((r) => r[1] === maDonHang)
      .map((r, i) => ({
        stt: i + 1,
        tenSanPham: r[8],
        soLuong: r[14],
        donVi: r[13],
        tongSoLuong: r[15],
        ghiChu: "",
      }));

    console.log(`✔️ Tìm thấy ${products.length} sản phẩm.`);

    // 4. Logo & watermark
    const logoBase64 = await loadDriveImageBase64(LOGO_FILE_ID);
    const watermarkBase64 = await loadDriveImageBase64(WATERMARK_FILE_ID);

    // 5. Render ngay
    res.render("bbntnk", {
      donHang,
      products,
      logoBase64,
      watermarkBase64,
      autoPrint: true,
      maDonHang,
      pathToFile: "",
    });

    // 6. Gọi AppScript lưu HTML
    (async () => {
      try {
        const renderedHtml = await renderFileAsync(
          path.join(__dirname, "views", "bbntnk.ejs"),
          {
            donHang,
            products,
            logoBase64,
            watermarkBase64,
            autoPrint: false,
            maDonHang,
            pathToFile: "",
          }
        );

        const GAS_WEBAPP_URL_BBNTNK = process.env.GAS_WEBAPP_URL_BBNTNK;
        if (GAS_WEBAPP_URL_BBNTNK) {
          const resp = await fetch(GAS_WEBAPP_URL_BBNTNK, {
            method: "POST",
            headers: { "Content-Type": "application/x-www-form-urlencoded" },
            body: new URLSearchParams({
              orderCode: maDonHang,
              html: renderedHtml,
            }),
          });

          const data = await resp.json();
          console.log("✔️ AppScript trả về:", data);

          const pathToFile = data.pathToFile || `BBNTNK/${data.fileName}`;
          await sheets.spreadsheets.values.update({
            spreadsheetId: SPREADSHEET_ID,
            range: `file_BBNT_ct!D${lastRowWithData}`,
            valueInputOption: "RAW",
            requestBody: { values: [[pathToFile]] },
          });
          console.log("✔️ Đã ghi đường dẫn:", pathToFile);
        }
      } catch (err) {
        console.error("❌ Lỗi gọi AppScript BBNTNK:", err);
      }
    })();
  } catch (err) {
    console.error("❌ Lỗi khi xuất BBNTNK:", err.stack || err.message);
    res.status(500).send("Lỗi server: " + (err.message || err));
  }
});



app.get("/gghnk", async (req, res) => {
    try {
        console.log("▶️ Bắt đầu xuất GGHNK ...");

        // --- Lấy mã đơn hàng ---
        const gghRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: "File_GGH_ct!B:B",
        });
        const colB = gghRes.data.values ? gghRes.data.values.flat() : [];
        const lastRowWithData = colB.length;
        const maDonHang = colB[lastRowWithData - 1];
        if (!maDonHang)
            return res.send("⚠️ Không tìm thấy dữ liệu ở cột B sheet File_GGH_ct.");

        console.log(`✔️ Mã đơn hàng: ${maDonHang} (dòng ${lastRowWithData})`);

        // --- Lấy đơn hàng ---
        const donHangRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: "Don_hang!A1:BJ",
        });
        const rows = donHangRes.data.values || [];
        const data = rows.slice(1);
        const donHang =
            data.find((r) => r[5] === maDonHang) ||
            data.find((r) => r[6] === maDonHang);
        if (!donHang)
            return res.send("❌ Không tìm thấy đơn hàng với mã: " + maDonHang);

        // --- Logo ---
        const logoBase64 = await loadDriveImageBase64(LOGO_FILE_ID);

        // --- Render ngay cho client ---
        res.render("gghnk", {
            donHang,
            logoBase64,
            autoPrint: false,
            maDonHang,
            pathToFile: ""
        });

        // --- Sau khi render xong thì gọi AppScript ngầm ---
        (async () => {
            try {
                const renderedHtml = await renderFileAsync(
                    path.join(__dirname, "views", "gghnk.ejs"),
                    {
                        donHang,
                        logoBase64,
                        autoPrint: false,
                        maDonHang,
                        pathToFile: ""
                    }
                );

                // Gọi GAS webapp tương ứng (cần thêm biến môi trường GAS_WEBAPP_URL_GGH)
                const GAS_WEBAPP_URL_GGH = process.env.GAS_WEBAPP_URL_GGH;
                if (GAS_WEBAPP_URL_GGH) {
                    const resp = await fetch(GAS_WEBAPP_URL_GGH, {
                        method: "POST",
                        headers: { "Content-Type": "application/x-www-form-urlencoded" },
                        body: new URLSearchParams({
                            orderCode: maDonHang,
                            html: renderedHtml
                        })
                    });

                    const data = await resp.json();
                    console.log("✔️ AppScript trả về:", data);

                    const pathToFile = data.pathToFile || `GGH/${data.fileName}`;
                    await sheets.spreadsheets.values.update({
                        spreadsheetId: SPREADSHEET_ID,
                        range: `File_GGH_ct!D${lastRowWithData}`,
                        valueInputOption: "RAW",
                        requestBody: { values: [[pathToFile]] },
                    });
                    console.log("✔️ Đã ghi đường dẫn:", pathToFile);
                } else {
                    console.log("⚠️ Chưa cấu hình GAS_WEBAPP_URL_GGH");
                }

            } catch (err) {
                console.error("❌ Lỗi gọi AppScript:", err);
            }
        })();

    } catch (err) {
        console.error("❌ Lỗi khi xuất GGH:", err.stack || err.message);
        res.status(500).send("Lỗi server: " + (err.message || err));
    }
});


// --- Route /dntu-<ma> ---
app.get("/dntu-:ma", async (req, res) => {
  try {
    const maTamUng = req.params.ma;
    console.log("▶️ Xuất giấy đề nghị tạm ứng:", maTamUng);

    // Lấy dữ liệu sheet
    const resp = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_HC_ID,
      range: "data_tam_ung_thanh_toan!A:Z",
    });
    const rows = resp.data.values || [];
    const data = rows.slice(1);

    // Tìm dòng có cột H (index 7) == maTamUng
    const record = data.find((r) => r[7] === maTamUng);
    if (!record) {
      return res.send("❌ Không tìm thấy mã tạm ứng: " + maTamUng);
    }

    // Map dữ liệu theo form
    const formData = {
      maTamUng: record[7],     // H
      ngayTamUng: formatVietnameseDate(record[4]),   // E
      ten: record[2],          // C
      boPhan: record[3],       // D
      soTien: formatNumber(record[9]),       // J
      soTienChu: numberToWords(record[9]),
      lyDo: record[8],         // I
      taikhoannhantu: record[11], //J
      thoiHan: record[12],     // M
    };

    // Logo + watermark
    const logoBase64 = await loadDriveImageBase64(LOGO_FILE_ID);
    const watermarkBase64 = await loadDriveImageBase64(WATERMARK_FILE_ID);

    // Render EJS
    res.render("dntu", {
      formData,
      logoBase64,
      watermarkBase64,
      autoPrint: true,
    });
  } catch (err) {
    console.error("❌ Lỗi DNTU:", err.stack || err.message);
    res.status(500).send("Lỗi server: " + (err.message || err));
  }
});

// --- Route /dnhu-<ma> ---
app.get("/dnhu-:ma", async (req, res) => {
  try {
    const maTamUng = req.params.ma;
    console.log("▶️ Xuất giấy đề nghị tạm ứng:", maTamUng);

    // Lấy dữ liệu sheet
    const resp = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_HC_ID,
      range: "data_tam_ung_thanh_toan!A:AF",
    });
    const rows = resp.data.values || [];
    const data = rows.slice(1);

    // Tìm dòng có cột H (index 7) == maTamUng
    const record = data.find((r) => r[7] === maTamUng);
    if (!record) {
      return res.send("❌ Không tìm thấy mã tạm ứng: " + maTamUng);
    }

    // Map dữ liệu theo form
    const formData = {
      maTamUng: record[7],     // H
      ngayhoanUng: formatVietnameseDate(record[23]),   // E
      ten: record[27],          // C
      boPhan: record[3],       // D
      soTien: formatNumber(record[9]),       // J
      soTienChu: numberToWords(record[9]),
      lyDo: record[22],         // I
    };

    // Logo + watermark
    const logoBase64 = await loadDriveImageBase64(LOGO_FILE_ID);
    const watermarkBase64 = await loadDriveImageBase64(WATERMARK_FILE_ID);

    // Render EJS
    res.render("dnhu", {
      formData,
      logoBase64,
      watermarkBase64,
      autoPrint: true,
    });
  } catch (err) {
    console.error("❌ Lỗi DNHU:", err.stack || err.message);
    res.status(500).send("Lỗi server: " + (err.message || err));
  }
});


// --- Route /dntt-<ma> ---
app.get("/dntt-:ma", async (req, res) => {
  try {
    const maTamUng = req.params.ma;
    console.log("▶️ Xuất giấy đề nghị thanh toán:", maTamUng);

    // Lấy dữ liệu sheet
    const resp = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_HC_ID,
      range: "data_tam_ung_thanh_toan!A:AF",
    });
    const rows = resp.data.values || [];
    const data = rows.slice(1);

    // Tìm dòng có cột H (index 7) == maTamUng
    const record = data.find((r) => r[7] === maTamUng);
    if (!record) {
      return res.send("❌ Không tìm thấy mã tạm ứng: " + maTamUng);
    }

    // Map dữ liệu theo form
    const formData = {
      maTamUng: record[7],     // H
      ngayhoanUng: formatVietnameseDate(record[23]),   // E
      ten: record[27],          // C
      boPhan: record[3],       // D
      soTientu: formatNumber(record[9]),       // J
      soTientuChu: numberToWords(record[9]),
      soTienthucchi: formatNumber(record[24]),       // J
      soTienthucchiChu: numberToWords(record[24]),
      soTienthanhtoan: formatNumber(record[29]),       // J
      soTienthanhtoanChu: numberToWords(record[29]),
      lyDo: record[22],        // I
      sotknhantien: record[28],
    };

    // Logo + watermark
    const logoBase64 = await loadDriveImageBase64(LOGO_FILE_ID);
    const watermarkBase64 = await loadDriveImageBase64(WATERMARK_FILE_ID);

    // Render EJS
    res.render("dntt", {
      formData,
      logoBase64,
      watermarkBase64,
      autoPrint: true,
    });
  } catch (err) {
    console.error("❌ Lỗi DNTT:", err.stack || err.message);
    res.status(500).send("Lỗi server: " + (err.message || err));
  }
});

app.use(express.static(path.join(__dirname, 'public')));
// --- Debug ---
app.get("/debug", (_req, res) => {
    res.json({ spreadsheetId: SPREADSHEET_ID, clientEmail: credentials.client_email, gasWebappUrl: GAS_WEBAPP_URL });
});

// --- Start server ---
app.listen(PORT, () => console.log(`✅ Server is running on port ${PORT}`));


// Hàm chuyển số thành chữ (thêm vào app.js)
function numberToWords(number) {
            const units = ['', 'một', 'hai', 'ba', 'bốn', 'năm', 'sáu', 'bảy', 'tám', 'chín'];
            const positions = ['', 'nghìn', 'triệu', 'tỷ', 'nghìn tỷ', 'triệu tỷ'];
            
            if (number === 0) return 'không';
            
            let words = '';
            let position = 0;
            
            do {
                const block = number % 1000;
                if (block !== 0) {
                    words = readBlock(block) + (position > 0 ? ' ' + positions[position] + ' ' : '') + words;
                }
                position++;
                number = Math.floor(number / 1000);
            } while (number > 0);
            
            return words.trim() + ' đồng';
            
            function readBlock(number) {
                let str = '';
                const hundreds = Math.floor(number / 100);
                const tens = Math.floor((number % 100) / 10);
                const ones = number % 10;
                
                if (hundreds > 0) {
                    str += units[hundreds] + ' trăm ';
                }
                
                if (tens === 0) {
                    if (ones > 0 && hundreds > 0) {
                        str += 'lẻ ';
                    }
                } else if (tens === 1) {
                    str += 'mười ';
                } else {
                    str += units[tens] + ' mươi ';
                }
                
                if (ones > 0) {
                    if (tens > 1 && ones === 1) {
                        str += 'mốt';
                    } else if (tens > 0 && ones === 5) {
                        str += 'lăm';
                    } else {
                        str += units[ones];
                    }
                }
                
                return str;
            }
}


function formatNumber(num) {
  if (!num) return "0";
  return num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");
}

// Hàm chuyển định dạng ngày tháng năm
function formatVietnameseDate(dateStr) {
      try {
        const d = new Date(dateStr);
        if (isNaN(d)) return dateStr; // Nếu không parse được thì trả nguyên
        const day = ("0" + d.getDate()).slice(-2);
        const month = ("0" + (d.getMonth() + 1)).slice(-2);
        const year = d.getFullYear();
        return `Ngày ${day} tháng ${month} năm ${year}`;
      } catch (e) {
        return dateStr;
      }
    }
