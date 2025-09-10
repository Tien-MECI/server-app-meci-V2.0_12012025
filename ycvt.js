import { google } from 'googleapis';

console.log('🚀 Đang load module ycvt.js...');

async function prepareYcvtData(auth, spreadsheetId, spreadsheetHcId) {
  console.log('▶️ Bắt đầu chuẩn bị dữ liệu cho YCVT...');
  const sheets = google.sheets({ version: 'v4', auth });

  // Hàm hỗ trợ: paste → đọc → clear
  async function pasteAndRead(rowIndex, targetValues) {
    const pasteRange = `Data_bom!F${rowIndex + 1}:L${rowIndex + 1}`;
    console.log(`📌 Paste targetValues vào ${pasteRange}`, targetValues);

    // Paste dữ liệu
    await sheets.spreadsheets.values.update({
      spreadsheetId: spreadsheetHcId,
      range: pasteRange,
      valueInputOption: 'RAW',
      requestBody: { values: [targetValues] }
    });

    // Đọc lại kết quả (Google Sheets đã tính công thức liên quan)
    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: spreadsheetHcId,
      range: `Data_bom!A${rowIndex + 1}:N${rowIndex + 1}`
    });
    const rowWithCalculated = result.data.values[0] || [];

    // Clear lại vùng vừa paste để không ghi đè công thức gốc
    await sheets.spreadsheets.values.clear({
      spreadsheetId: spreadsheetHcId,
      range: pasteRange
    });

    return rowWithCalculated;
  }

  try {
    // Load dữ liệu
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

    console.log(`✔️ Đã lấy dữ liệu: ${data1.length} rows (Don_hang_PVC_ct), ${data2.length} rows (Don_hang), ${data3.length} rows (Data_bom), ${data5.length} rows (File_BOM_ct)`);

    // Lấy mã đơn hàng cuối cùng ở File_BOM_ct
    const colB = data5.map(row => row[1]).filter(v => v);
    const lastRowWithData = colB.length;
    const d4Value = colB[lastRowWithData - 1];
    if (!d4Value) throw new Error('Không tìm thấy mã đơn hàng trong File_BOM_ct');
    console.log(`✔️ Mã đơn hàng: ${d4Value} (dòng ${lastRowWithData})`);

    const donHang = data2.slice(1).find(row => row[5] === d4Value || row[6] === d4Value);
    if (!donHang) throw new Error(`Không tìm thấy đơn hàng với mã: ${d4Value}`);

    // Tập hợp hValues từ Don_hang_PVC_ct
    const hValues = data1.slice(1)
      .filter(row => row[1] === d4Value)
      .map((row, i) => ({
        stt: i + 1,
        hValue: row[7] || '',
        rowData: row
      }));
    console.log(`✔️ Tìm thấy ${hValues.length} sản phẩm với hValue.`);

    const columnsToCopyBase = [17, 18, 19, 20, 21, 22, 23, 24, 29]; // Cột từ Don_hang_PVC_ct
    let tableData = [];

    // Xử lý từng hValue
    for (const hObj of hValues) {
      const hValue = hObj.hValue;

      // Tìm tất cả row ở Data_bom có cột C = hValue
      const matchingIndexes = data3
        .map((row, idx) => ({ row, idx }))
        .filter(item => item.row[2] === hValue);

      if (matchingIndexes.length > 0) {
        let isMainRowProcessed = false;

        for (const { row, idx } of matchingIndexes) {
          let rowData;

          if (!isMainRowProcessed) {
            // Paste targetValues vào F:L
            const targetValues = columnsToCopyBase.map(i => hObj.rowData[i - 1] || '');
            rowData = await pasteAndRead(idx, targetValues);
            isMainRowProcessed = true;
          } else {
            // Với row phụ, chỉ đọc lại giá trị tính sẵn
            rowData = row;
          }

          tableData.push({
            stt: hObj.stt,
            row: rowData.slice(1, 14) // Lấy B:N
          });
          console.log(`✔️ Đã xử lý row Data_bom ${idx + 1} cho hValue ${hValue}`);
        }
      } else {
        console.warn(`⚠️ Không tìm thấy hValue ${hValue} trong Data_bom cột C`);
      }
    }

    console.log('📋 tableData:', JSON.stringify(tableData, null, 2));

    // Các thông tin từ Don_hang
    const matchingRows = data2.slice(1).filter(row => row[5] === d4Value || row[6] === d4Value);
    const l4Value = matchingRows[0] ? (matchingRows[0][8] || '') : '';
    const d5Values = matchingRows.flatMap(row => row[83] || []).filter(v => v).join(', ');
    const h5Values = matchingRows.flatMap(row => row[36] || []).filter(v => v).join(', ');
    const h6Values = matchingRows.flatMap(row => row[37] || []).filter(v => v).join(', ');
    const d6Values = matchingRows
      .flatMap(row => row[48] ? new Date(row[48]).toLocaleDateString('vi-VN', { day: '2-digit', month: '2-digit', year: 'numeric' }) : [])
      .filter(v => v)
      .join('<br>');

    // Chuẩn bị bảng tổng hợp
    const uniqueB = [...new Set(tableData.map(item => item.row[1]).filter(v => v && v !== 'Mã SP' && v !== 'Mã vật tư sản xuất'))];
    const uniqueC = [...new Set(tableData.map(item => item.row[2]).filter(v => v && v !== 'Mã vật tư xuất kèm' && v !== 'Mã vật tư sản xuất'))];

    const summaryDataB = uniqueB.map((b, i) => {
      const sum = tableData
        .filter(item => item.row[1] === b || item.row[2] === b)
        .reduce((sum, item) => sum + (parseFloat((item.row[8] || '').toString().replace(',', '.')) || 0), 0);
      const desc = tableData.find(item => item.row[1] === b || item.row[2] === b)?.row[3] || '';
      return { stt: i + 1, b, sum, desc };
    });
    const summaryDataC = uniqueC.map((c, i) => {
      const sum = tableData
        .filter(item => item.row[1] === c || item.row[2] === c)
        .reduce((sum, item) => sum + (parseFloat((item.row[10] || '').toString().replace(',', '.')) || 0), 0);
      const desc = tableData.find(item => item.row[1] === c || item.row[2] === c)?.row[3] || '';
      return { stt: summaryDataB.length + i + 1, c, sum, desc };
    });

    console.log(`✔️ Tạo ${summaryDataB.length} mục B và ${summaryDataC.length} mục C trong bảng tổng hợp.`);

    // Kiểm tra cột E/I/J có dữ liệu hay không
    const hasDataE = tableData.some(item => item.row[4] && item.row[4].toString().trim() !== '');
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
      hasDataE,
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
