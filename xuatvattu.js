// ycvt.js
import { google } from 'googleapis';

console.log('🚀 Đang load module ycvt.js...');

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * preparexkvtData
 * - auth: OAuth2 client
 * - spreadsheetId: id của workbook chính (chứa Don_hang_PVC_ct)
 * - spreadsheetHcId: id workbook chứa Data_bom
 * - spreadsheetKhvtId: id workbook chứa xuat_kho_VT
 * - maDonHang: mã đơn hàng được cung cấp
 */
async function preparexkvtData(auth, spreadsheetId, spreadsheetHcId, spreadsheetKhvtId, maDonHang) {
  console.log('▶️ Bắt đầu preparexkvtData...');
  const sheets = google.sheets({ version: 'v4', auth });

  async function batchPaste(spreadsheetId, valueRanges) {
    if (!valueRanges || valueRanges.length === 0) return;
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: { valueInputOption: 'RAW', data: valueRanges }
    });
  }

  async function batchClear(spreadsheetId, ranges) {
    if (!ranges || ranges.length === 0) return;
    await sheets.spreadsheets.values.batchClear({
      spreadsheetId,
      requestBody: { ranges }
    });
  }

  try {
    // 1) Lấy dữ liệu ban đầu
    const [data1Res, data3Res] = await Promise.all([
      sheets.spreadsheets.values.get({ spreadsheetId, range: 'Don_hang_PVC_ct!A1:AE' }),
      sheets.spreadsheets.values.get({ spreadsheetId: spreadsheetHcId, range: 'Data_bom!A1:O' })
    ]);

    const data1 = data1Res.data.values || [];
    const data3 = data3Res.data.values || [];

    console.log(`✔️ Lấy dữ liệu xong: Don_hang_PVC_ct=${data1.length}, Data_bom=${data3.length}`);

    // 2) Kiểm tra maDonHang
    if (!maDonHang) throw new Error('Không có mã đơn hàng được cung cấp');
    console.log(`✔️ Mã đơn hàng: ${maDonHang}`);

    // 3) Tìm hValues từ Don_hang_PVC_ct
    const hValues = data1.slice(1)
      .map((r, idx) => ({ row: r, idx }))
      .filter(o => String(o.row[1] || '').trim() === String(maDonHang).trim())
      .map((o, i) => ({ stt: i + 1, hValue: o.row[7] || '', rowData: o.row }));

    console.log(`✔️ Tìm thấy ${hValues.length} hValue trong Don_hang_PVC_ct`);

    // 4) Chuẩn bị maps để tổng hợp dữ liệu
    const sanPhamMap = new Map(); // key: C (mã sản phẩm), value: { sumK: number, L: string }
    const vatTuMap = new Map(); // key: D (mã vật tư), value: { sumL: number, M: string }

    // 5) Xử lý tuần tự từng hValue
    const columnsToCopyBase = [17, 18, 19, 20, 21, 22, 23, 24, 29]; // 9 cột

    for (const hObj of hValues) {
      const hValue = hObj.hValue;
      if (!hValue) continue;

      // Tìm tất cả row trong Data_bom có C == hValue
      const matchesC = [];
      for (let i = 0; i < data3.length; i++) {
        const row = data3[i] || [];
        if (String(row[2] || '').trim() === String(hValue).trim()) matchesC.push(i);
      }
      if (matchesC.length === 0) continue;

      // Build dữ liệu để paste
      const targetValues = columnsToCopyBase.map(colIndex =>
        hObj.rowData[colIndex - 1] !== undefined ? hObj.rowData[colIndex - 1] : ''
      );

      const pasteValueRanges = [];
      const pastedRanges = [];
      for (const idx of matchesC) {
        const rowNum = idx + 1;
        const range = `Data_bom!F${rowNum}:N${rowNum}`;
        pasteValueRanges.push({ range, values: [targetValues] });
        pastedRanges.push(range);
      }

      // Paste riêng cho hValue này
      if (pasteValueRanges.length > 0) {
        console.log(`📥 Paste ${hValue}: ${pasteValueRanges.length} ranges...`);
        await batchPaste(spreadsheetHcId, pasteValueRanges);
        await sleep(600);

        // Đọc lại Data_bom (đến O)
        let updatedData3 = null;
        for (let attempts = 0; attempts < 5; attempts++) {
          const res = await sheets.spreadsheets.values.get({
            spreadsheetId: spreadsheetHcId,
            range: 'Data_bom!A1:O'
          });
          updatedData3 = res.data.values || [];
          const someBpopulated = updatedData3.some(r => r && r[1] && String(r[1]).trim() !== '');
          if (someBpopulated) break;
          await sleep(600);
        }

        // Thu thập và tổng hợp dữ liệu theo A == hValue và dựa trên O
        for (const row of updatedData3) {
          if (String(row?.[0] || '').trim() === String(hValue).trim()) {
            const oValue = (row[14] || '').trim(); // Cột O (index 14)
            if (oValue === 'Sản phẩm') {
              const c = (row[2] || '').trim(); // Cột C (index 2)
              const k_value = (row[10] || '').toString().trim(); // Cột K (index 10) - số lượng
              const k_clean = k_value.replace(/\./g, '').replace(',', '.'); // Xử lý định dạng số Việt Nam
              const k = parseFloat(k_clean) || 0;
              const l = (row[11] || '').trim(); // Cột L (index 11)
              if (c) {
                console.log(`Sản phẩm (hValue=${hValue}): C=${c}, K_value=${k_value}, k=${k}, L=${l}`);
                if (sanPhamMap.has(c)) {
                  const obj = sanPhamMap.get(c);
                  obj.sumK += k;
                } else {
                  sanPhamMap.set(c, { sumK: k, L: l });
                }
              }
            } else if (oValue === 'Vật tư') {
              const d = (row[3] || '').trim(); // Cột D (index 3)
              const l_value = (row[11] || '').toString().trim(); // Cột L (index 11) - số lượng
              const l_clean = l_value.replace(/\./g, '').replace(',', '.'); // Xử lý định dạng số Việt Nam
              const l_sum = parseFloat(l_clean) || 0;
              const m = (row[12] || '').trim(); // Cột M (index 12)
              if (d) {
                console.log(`Vật tư (hValue=${hValue}): D=${d}, L_value=${l_value}, l_sum=${l_sum}, M=${m}`);
                if (vatTuMap.has(d)) {
                  const obj = vatTuMap.get(d);
                  obj.sumL += l_sum;
                } else {
                  vatTuMap.set(d, { sumL: l_sum, M: m });
                }
              }
            }
          }
        }

        // Clear lại
        if (pastedRanges.length > 0) {
          console.log(`🧹 Clear ${hValue}: ${pastedRanges.length} ranges...`);
          await batchClear(spreadsheetHcId, pastedRanges);
        }
      }
    }

    // Log để kiểm tra maps
    console.log('✔️ sanPhamMap:', Array.from(sanPhamMap.entries()));
    console.log('✔️ vatTuMap:', Array.from(vatTuMap.entries()));

    // 6) Lấy dữ liệu hiện tại từ sheet xuat_kho_VT để tìm hàng cuối cùng
    const xuatDataRes = await sheets.spreadsheets.values.get({
      spreadsheetId: spreadsheetKhvtId,
      range: 'xuat_kho_VT!A1:G'
    });
    const xuatData = xuatDataRes.data.values || [];
    let lastRow = xuatData.length + 1; // Hàng tiếp theo để dán

    // 7) Chuẩn bị dữ liệu để dán vào xuat_kho_VT
    const valueRanges = [];
    const currentDate = new Date().toLocaleDateString('vi-VN', { day: '2-digit', month: '2-digit', year: 'numeric' });

    // Xử lý sanPhamMap
    for (const [code, { sumK, L }] of sanPhamMap) {
      const uniqueId = Math.random().toString(36).substring(2, 10).toUpperCase(); // Unique ID 8 ký tự
      const values = [
        uniqueId,      // A: Unique ID
        maDonHang,     // B: Mã đơn hàng
        currentDate,   // C: dd/mm/yyyy
        code,          // D: Mã (từ C hoặc D)
        '',            // E: Rỗng
        sumK,          // F: Sum số lượng
        L              // G: Đơn vị tính
      ];
      valueRanges.push({
        range: `xuat_kho_VT!A${lastRow}:G${lastRow}`,
        values: [values]
      });
      lastRow++;
    }

    // Xử lý vatTuMap
    for (const [code, { sumL, M }] of vatTuMap) {
      const uniqueId = Math.random().toString(36).substring(2, 10).toUpperCase(); // Unique ID 8 ký tự
      const values = [
        uniqueId,      // A: Unique ID
        maDonHang,     // B: Mã đơn hàng
        currentDate,   // C: dd/mm/yyyy
        code,          // D: Mã (từ C hoặc D)
        '',            // E: Rỗng
        sumL,          // F: Sum số lượng
        M              // G: Đơn vị tính
      ];
      valueRanges.push({
        range: `xuat_kho_VT!A${lastRow}:G${lastRow}`,
        values: [values]
      });
      lastRow++;
    }

    // 8) Paste vào xuat_kho_VT
    if (valueRanges.length > 0) {
      console.log(`📥 Paste vào xuat_kho_VT: ${valueRanges.length} rows...`);
      await batchPaste(spreadsheetKhvtId, valueRanges);
    }

    return { success: true, message: 'Xử lý và dán dữ liệu thành công' };

  } catch (err) {
    console.error('❌ Lỗi trong preparexkvtData:', err.stack || err.message);
    throw err;
  }
}

export { preparexkvtData };