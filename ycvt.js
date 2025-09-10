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

  // Hỗ trợ: batch paste nhiều range cùng lúc
  async function batchPaste(valueRanges) {
    if (!valueRanges || valueRanges.length === 0) return;
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: spreadsheetHcId,
      requestBody: {
        valueInputOption: 'RAW',
        data: valueRanges
      }
    });
  }

  // Hỗ trợ: batch clear nhiều range cùng lúc
  async function batchClear(ranges) {
    if (!ranges || ranges.length === 0) return;
    await sheets.spreadsheets.values.batchClear({
      spreadsheetId: spreadsheetHcId,
      requestBody: { ranges }
    });
  }

  try {
    // 1) Lấy dữ liệu ban đầu (1 lần)
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

    // 2) Tìm d4Value = last non-empty in File_BOM_ct col B
    let d4Value = '';
    let lastRowWithData = 0;
    for (let i = data5.length - 1; i >= 0; i--) {
      const v = (data5[i] && data5[i][1]) ? String(data5[i][1]).trim() : '';
      if (v !== '') { d4Value = v; lastRowWithData = i + 1; break; }
    }
    if (!d4Value) throw new Error('Không tìm thấy mã đơn hàng trong File_BOM_ct cột B');
    console.log(`✔️ Mã đơn hàng: ${d4Value} (dòng ${lastRowWithData})`);

    // 3) Lấy hValues từ Don_hang_PVC_ct: các row có col B == d4Value; lấy H (index 7)
    const hValues = data1.slice(1)
      .map((r, idx) => ({ row: r, idx }))
      .filter(o => String(o.row[1] || '').trim() === String(d4Value).trim())
      .map((o, i) => ({ stt: i + 1, hValue: o.row[7] || '', rowData: o.row }));

    console.log(`✔️ Tìm thấy ${hValues.length} hValue trong Don_hang_PVC_ct`);

    // 4) columnsToCopyBase (từ Don_hang_PVC_ct) -> paste vào F:N (9 cột)
    const columnsToCopyBase = [17, 18, 19, 20, 21, 22, 23, 24, 29]; // 1-based indices
    const pasteValueRanges = []; // dùng cho batchUpdate
    const pastedRanges = [];     // danh sách ranges đã paste (để clear sau đó)

    // Tạo valueRanges: cho mỗi hValue tìm các hàng có C == hValue -> paste vào F:N trên chính những hàng đó
    for (const hObj of hValues) {
      const hValue = hObj.hValue;
      if (!hValue) {
        console.warn('⚠️ hValue trống, bỏ qua');
        continue;
      }

      // tìm tất cả index i trong data3 có col C (index 2) === hValue
      const matchesC = [];
      for (let i = 0; i < data3.length; i++) {
        const row = data3[i] || [];
        if (String(row[2] || '').trim() === String(hValue).trim()) matchesC.push(i);
      }

      if (matchesC.length === 0) {
        console.log(`ℹ️ Không có hàng C === ${hValue} (Data_bom)`); 
        continue;
      }

      // targetValues lấy từ hObj.rowData
      const targetValues = columnsToCopyBase.map(colIndex => (hObj.rowData[colIndex - 1] !== undefined ? hObj.rowData[colIndex - 1] : ''));

      // push valueRanges cho mỗi hàng match
      for (const idx of matchesC) {
        const rowNum = idx + 1; // spreadsheet 1-based
        const range = `Data_bom!F${rowNum}:N${rowNum}`; // F..N (9 cột)
        pasteValueRanges.push({ range, values: [targetValues] });
        pastedRanges.push(range);
        console.log(`→ Will paste for hValue=${hValue} at row ${rowNum} range ${range}`);
      }
    }

    // 5) Nếu có range cần paste thì batch paste 1 lần
    if (pasteValueRanges.length > 0) {
      console.log(`📥 Batch paste ${pasteValueRanges.length} ranges into Data_bom (F:N) ...`);
      await batchPaste(pasteValueRanges);

      // chờ ngắn để Google Sheets tính (tùy tốc độ bạn có thể tăng)
      const WAIT_MS = 600;
      await sleep(WAIT_MS);

      // optional: có thể poll thêm vài lần nếu công thức nặng — ở đây ta làm 3 attempts nhỏ
      let attempts = 0;
      const MAX_ATTEMPTS = 5;
      let updatedData3 = null;
      while (attempts < MAX_ATTEMPTS) {
        // 6) Đọc lại toàn bộ Data_bom!A:N 1 lần (để có giá trị đã được tính)
        const res = await sheets.spreadsheets.values.get({
          spreadsheetId: spreadsheetHcId,
          range: 'Data_bom!A1:N'
        });
        updatedData3 = res.data.values || [];
        // Heuristic: nếu updatedData3 length >= original length, và có ít nhất 1 hàng có B (index 1) khác '' thì break
        const someBpopulated = updatedData3.some(r => (r && r[1] && String(r[1]).trim() !== ''));
        if (someBpopulated || attempts === MAX_ATTEMPTS - 1) {
          console.log(`📖 Đã đọc Data_bom (attempt ${attempts + 1}) — rows: ${updatedData3.length}`);
          break;
        }
        attempts++;
        console.log(`⏳ Chưa có dữ liệu tính xong, đợi thêm ${WAIT_MS}ms (attempt ${attempts})`);
        await sleep(WAIT_MS);
      }

      // 7) Từ updatedData3, lấy B:N cho các hàng có A === hValue
      const tableData = [];
      for (const hObj of hValues) {
        const hValue = hObj.hValue;
        if (!hValue) continue;
        // find all rows where col A (index 0) === hValue
        for (let i = 0; i < updatedData3.length; i++) {
          const row = updatedData3[i] || [];
          if (String(row[0] || '').trim() === String(hValue).trim()) {
            // take B:N -> slice(1,14)
            const sliceBN = row.slice(1, 14);
            // normalize to length 13
            while (sliceBN.length < 13) sliceBN.push('');
            tableData.push({ stt: hObj.stt, row: sliceBN });
          }
        }
      }

      console.log(`✔️ Đã thu thập tableData từ updated Data_bom (dựa trên A==hValue): ${tableData.length} rows`);

      // 8) Clear lại tất cả pastedRanges (batch)
      console.log(`🧹 Clear ${pastedRanges.length} pasted ranges (F:N) ...`);
      await batchClear(pastedRanges);

      // 9) Tiếp tục build summary / trả về
      const uniqueB = [...new Set(tableData.map(item => item.row[1]).filter(v => v && v !== 'Mã SP' && v !== 'Mã vật tư sản xuất'))];
      const uniqueC = [...new Set(tableData.map(item => item.row[2]).filter(v => v && v !== 'Mã vật tư xuất kèm' && v !== 'Mã vật tư sản xuất'))];

      const summaryDataB = uniqueB.map((b, i) => {
        const sum = tableData
          .filter(item => item.row[1] === b || item.row[2] === b)
          .reduce((s, item) => s + (parseFloat((item.row[8] || '').toString().replace(',', '.')) || 0), 0);
        const desc = tableData.find(item => item.row[1] === b || item.row[2] === b)?.row[3] || '';
        return { stt: i + 1, code: b, sum, desc };
      });
      const summaryDataC = uniqueC.map((c, i) => {
        const sum = tableData
          .filter(item => item.row[1] === c || item.row[2] === c)
          .reduce((s, item) => s + (parseFloat((item.row[10] || '').toString().replace(',', '.')) || 0), 0);
        const desc = tableData.find(item => item.row[1] === c || item.row[2] === c)?.row[3] || '';
        return { stt: summaryDataB.length + i + 1, code: c, sum, desc };
      });

      // thông tin Don_hang
      const matchingRows = data2.slice(1).filter(row => String(row[5] || '').trim() === String(d4Value).trim() || String(row[6] || '').trim() === String(d4Value).trim());
      const l4Value = matchingRows[0] ? (matchingRows[0][8] || '') : '';
      const d5Values = matchingRows.map(r => r[83]).filter(v => v).join(', ');
      const h5Values = matchingRows.map(r => r[36]).filter(v => v).join(', ');
      const h6Values = matchingRows.map(r => r[37]).filter(v => v).join(', ');
      const d6Values = matchingRows
        .map(r => r[48] ? new Date(r[48]).toLocaleDateString('vi-VN', { day: '2-digit', month: '2-digit', year: 'numeric' }) : '')
        .filter(v => v)
        .join('<br>');

      // Kiểm tra cột F (index 5 trong B:N)
      const hasDataF = tableDataFrom7.some(item => item.row[5] && item.row[5].toString().trim() !== '');

      // Kiểm tra cột I (index 8 trong B:N)
      const hasDataI = tableDataFrom7.some(item => item.row[8] && item.row[8].toString().trim() !== '');

      // Kiểm tra cột J (index 9 trong B:N)
      const hasDataJ = tableDataFrom7.some(item => item.row[9] && item.row[9].toString().trim() !== '');


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

    } else {
      // Nếu không có range cần paste (không tìm thấy C==hValue nào), vẫn phải tạo tableData bằng A==hValue trên data3 ban đầu
      console.log('ℹ️ Không có paste operations (không tìm thấy bất cứ C==hValue nào). Thu thập B:N dựa trên A==hValue từ data3 ban đầu.');

      const tableData = [];
      for (const hObj of hValues) {
        for (let i = 0; i < data3.length; i++) {
          const row = data3[i] || [];
          if (String(row[0] || '').trim() === String(hObj.hValue).trim()) {
            const sliceBN = row.slice(1, 14);
            while (sliceBN.length < 13) sliceBN.push('');
            tableData.push({ stt: hObj.stt, row: sliceBN });
          }
        }
      }

      const uniqueB = [...new Set(tableData.map(item => item.row[1]).filter(v => v && v !== 'Mã SP' && v !== 'Mã vật tư sản xuất'))];
      const uniqueC = [...new Set(tableData.map(item => item.row[2]).filter(v => v && v !== 'Mã vật tư xuất kèm' && v !== 'Mã vật tư sản xuất'))];

      const summaryDataB = uniqueB.map((b, i) => {
        const sum = tableData
          .filter(item => item.row[1] === b || item.row[2] === b)
          .reduce((s, item) => s + (parseFloat((item.row[8] || '').toString().replace(',', '.')) || 0), 0);
        const desc = tableData.find(item => item.row[1] === b || item.row[2] === b)?.row[3] || '';
        return { stt: i + 1, code: b, sum, desc };
      });
      const summaryDataC = uniqueC.map((c, i) => {
        const sum = tableData
          .filter(item => item.row[1] === c || item.row[2] === c)
          .reduce((s, item) => s + (parseFloat((item.row[10] || '').toString().replace(',', '.')) || 0), 0);
        const desc = tableData.find(item => item.row[1] === c || item.row[2] === c)?.row[3] || '';
        return { stt: summaryDataB.length + i + 1, code: c, sum, desc };
      });

      const matchingRows = data2.slice(1).filter(row => String(row[5] || '').trim() === String(d4Value).trim() || String(row[6] || '').trim() === String(d4Value).trim());
      const l4Value = matchingRows[0] ? (matchingRows[0][8] || '') : '';
      const d5Values = matchingRows.map(r => r[83]).filter(v => v).join(', ');
      const h5Values = matchingRows.map(r => r[36]).filter(v => v).join(', ');
      const h6Values = matchingRows.map(r => r[37]).filter(v => v).join(', ');
      const d6Values = matchingRows
        .map(r => r[48] ? new Date(r[48]).toLocaleDateString('vi-VN', { day: '2-digit', month: '2-digit', year: 'numeric' }) : '')
        .filter(v => v)
        .join('<br>');

      const hasDataF = tableData.some(item => item.row[5] && String(item.row[9]).trim() !== '');
      const hasDataI = tableData.some(item => item.row[8] && String(item.row[8]).trim() !== '');
      const hasDataJ = tableData.some(item => item.row[9] && String(item.row[5]).trim() !== '');

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
    }

  } catch (err) {
    console.error('❌ Lỗi trong prepareYcvtData:', err.stack || err.message);
    throw err;
  }
}

export { prepareYcvtData };
