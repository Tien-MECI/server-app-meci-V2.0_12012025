// ycvt.js
import { google } from 'googleapis';

console.log('🚀 Đang load module ycvt.js...');

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * prepareYcvtData
 * - auth: OAuth2 client
 * - spreadsheetId: id của workbook chính (chứa Don_hang_PVC_ct, Don_hang, File_BOM_ct)
 * - spreadsheetHcId: id workbook chứa Data_bom
 */
async function prepareYcvtData(auth, spreadsheetId, spreadsheetHcId) {
  console.log('▶️ Bắt đầu prepareYcvtData...');
  const sheets = google.sheets({ version: 'v4', auth });

  async function batchPaste(valueRanges) {
    if (!valueRanges || valueRanges.length === 0) return;
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: spreadsheetHcId,
      requestBody: { valueInputOption: 'RAW', data: valueRanges }
    });
  }

  async function batchClear(ranges) {
    if (!ranges || ranges.length === 0) return;
    await sheets.spreadsheets.values.batchClear({
      spreadsheetId: spreadsheetHcId,
      requestBody: { ranges }
    });
  }

  try {
    // 1) Lấy dữ liệu ban đầu
    const [data1Res, data2Res, data3Res, data5Res] = await Promise.all([
      sheets.spreadsheets.values.get({ spreadsheetId, range: 'Don_hang_PVC_ct!A1:AE' }),
      sheets.spreadsheets.values.get({ spreadsheetId, range: 'Don_hang!A1:CF' }),
      sheets.spreadsheets.values.get({ spreadsheetId: spreadsheetHcId, range: 'Data_bom!A1:N' }),
      sheets.spreadsheets.values.get({ spreadsheetId, range: 'File_BOM_ct!A1:D' })
    ]);

    const data1 = data1Res.data.values || [];
    const data2 = data2Res.data.values || [];
    const data3 = data3Res.data.values || [];
    const data5 = data5Res.data.values || [];

    console.log(`✔️ Lấy dữ liệu xong: Don_hang_PVC_ct=${data1.length}, Don_hang=${data2.length}, Data_bom=${data3.length}, File_BOM_ct=${data5.length}`);

    // 2) Lấy mã đơn hàng d4Value
    let d4Value = '';
    let lastRowWithData = 0;
    for (let i = data5.length - 1; i >= 0; i--) {
      const v = (data5[i] && data5[i][1]) ? String(data5[i][1]).trim() : '';
      if (v !== '') { d4Value = v; lastRowWithData = i + 1; break; }
    }
    if (!d4Value) throw new Error('Không tìm thấy mã đơn hàng trong File_BOM_ct cột B');
    console.log(`✔️ Mã đơn hàng: ${d4Value} (dòng ${lastRowWithData})`);

    // 3) Tìm hValues từ Don_hang_PVC_ct
    const hValues = data1.slice(1)
      .map((r, idx) => ({ row: r, idx }))
      .filter(o => String(o.row[1] || '').trim() === String(d4Value).trim())
      .map((o, i) => ({ stt: i + 1, hValue: o.row[7] || '', rowData: o.row }));

    console.log(`✔️ Tìm thấy ${hValues.length} hValue trong Don_hang_PVC_ct`);

    // 4) Chuẩn bị paste vào F:N
    const columnsToCopyBase = [17, 18, 19, 20, 21, 22, 23, 24, 29]; // 9 cột
    const pasteValueRanges = [];
    const pastedRanges = [];

    for (const hObj of hValues) {
      const hValue = hObj.hValue;
      if (!hValue) continue;

      // tìm tất cả row trong Data_bom có C == hValue
      const matchesC = [];
      for (let i = 0; i < data3.length; i++) {
        const row = data3[i] || [];
        if (String(row[2] || '').trim() === String(hValue).trim()) matchesC.push(i);
      }

      if (matchesC.length === 0) continue;

      const targetValues = columnsToCopyBase.map(colIndex =>
        hObj.rowData[colIndex - 1] !== undefined ? hObj.rowData[colIndex - 1] : ''
      );

      for (const idx of matchesC) {
        const rowNum = idx + 1;
        const range = `Data_bom!F${rowNum}:N${rowNum}`;
        pasteValueRanges.push({ range, values: [targetValues] });
        pastedRanges.push(range);
      }
    }

    // --------------------------
    // 5) Nếu có paste thì xử lý
    // --------------------------
    let tableData = [];
    if (pasteValueRanges.length > 0) {
      console.log(`📥 Batch paste ${pasteValueRanges.length} ranges (F:N) ...`);
      await batchPaste(pasteValueRanges);
      await sleep(600);

      let updatedData3 = null;
      for (let attempts = 0; attempts < 5; attempts++) {
        const res = await sheets.spreadsheets.values.get({
          spreadsheetId: spreadsheetHcId,
          range: 'Data_bom!A1:N'
        });
        updatedData3 = res.data.values || [];
        const someBpopulated = updatedData3.some(r => r && r[1] && String(r[1]).trim() !== '');
        if (someBpopulated) break;
        await sleep(600);
      }

      // Thu thập B:N
      for (const hObj of hValues) {
        for (const row of updatedData3) {
          if (String(row?.[0] || '').trim() === String(hObj.hValue).trim()) {
            const sliceBN = row.slice(1, 14);
            while (sliceBN.length < 13) sliceBN.push('');
            tableData.push({ stt: hObj.stt, row: sliceBN });
          }
        }
      }

      console.log(`✔️ Đã thu thập tableData từ Data_bom sau paste: ${tableData.length} rows`);

      // Clear lại
      if (pastedRanges.length > 0) {
        console.log(`🧹 Clear ${pastedRanges.length} ranges (F:N)...`);
        await batchClear(pastedRanges);
      }
    } else {
      // --------------------------
      // Không có paste
      // --------------------------
      console.log('ℹ️ Không có paste, lấy trực tiếp từ data3 ban đầu.');
      for (const hObj of hValues) {
        for (const row of data3) {
          if (String(row?.[0] || '').trim() === String(hObj.hValue).trim()) {
            const sliceBN = row.slice(1, 14);
            while (sliceBN.length < 13) sliceBN.push('');
            tableData.push({ stt: hObj.stt, row: sliceBN });
          }
        }
      }
    }

    // 6) Summary
    const uniqueB = [...new Set(tableData.map(item => item.row[1]).filter(v => v && v !== 'Mã SP' && v !== 'Mã vật tư sản xuất'))];
    const uniqueC = [...new Set(tableData.map(item => item.row[2]).filter(v => v && v !== 'Mã vật tư xuất kèm' && v !== 'Mã vật tư sản xuất'))];

    const summaryDataB = uniqueB.map((b, i) => {
      const relatedRows = tableData.filter(item => item.row[1] === b || item.row[2] === b);
      const sum = relatedRows.reduce((s, item) =>
        s + (parseFloat((item.row[9] || '').toString().replace(',', '.')) || 0), 0);
      const desc = relatedRows.find(item => item.row[3])?.row[3] || '';
      const DVT = relatedRows.find(item => item.row[10])?.row[10] || '';
      return { stt: i + 1, code: b, sum, desc, DVT };
    });

    const summaryDataC = uniqueC.map((c, i) => {
      const relatedRows = tableData.filter(item =>
        item.row[1] === c || item.row[2] === c
      );
      const sum = relatedRows.reduce((s, item) =>
        s + (parseFloat((item.row[9] || '').toString().replace(',', '.')) || 0), 0);
      const desc = relatedRows.find(item => item.row[3])?.row[3] || '';
      const DVT = relatedRows.find(item => item.row[10])?.row[10] || '';
      return { stt: summaryDataB.length + i + 1, code: c, sum, desc, DVT };
    });


    // 7) Thông tin Don_hang
    const matchingRows = data2.slice(1).filter(
      row => String(row[5] || '').trim() === String(d4Value).trim() || String(row[6] || '').trim() === String(d4Value).trim()
    );
    const l4Value = matchingRows[0] ? (matchingRows[0][8] || '') : '';
    const d5Values = matchingRows.map(r => r[83]).filter(v => v).join(', ');
    const h5Values = matchingRows.map(r => r[36]).filter(v => v).join(', ');
    const h6Values = matchingRows.map(r => r[37]).filter(v => v).join(', ');
    const d6Values = matchingRows
      .map(r => r[48] ? new Date(r[48]).toLocaleDateString('vi-VN', { day: '2-digit', month: '2-digit', year: 'numeric' }) : '')
      .filter(v => v).join('<br>');

    // 8) Flags kiểm tra dữ liệu
    const hasDataF = tableData.some(item => item.row[4] && item.row[4].toString().trim() !== '');
    const hasDataI = tableData.some(item => item.row[7] && item.row[7].toString().trim() !== '');
    const hasDataJ = tableData.some(item => item.row[8] && item.row[8].toString().trim() !== '');

    return {
      d4Value,
      l4Value,
      d3: new Date().toLocaleDateString('vi-VN', { day: '2-digit', month: '2-digit', year: 'numeric' }),
      d5Values,
      h5Values,
      h6Values,
      d6Values,
      tableData,
      summaryDataB,
      summaryDataC,
      hasDataF,
      hasDataI,
      hasDataJ,
      lastRowWithData
    };

  } catch (err) {
    console.error('❌ Lỗi trong prepareYcvtData:', err.stack || err.message);
    throw err;
  }
}

export { prepareYcvtData };
