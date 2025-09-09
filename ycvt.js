import { google } from 'googleapis';

console.log('🚀 Đang load module ycvt.js...');

async function prepareYcvtData(auth, spreadsheetId, spreadsheetHcId) {
    console.log('▶️ Bắt đầu chuẩn bị dữ liệu cho YCVT...');
    const sheets = google.sheets({ version: 'v4', auth });
    try {
        const [data1Res, data2Res, data3Res, data5Res] = await Promise.all([
            sheets.spreadsheets.values.get({ spreadsheetId, range: 'Don_hang_PVC_ct!A1:AE' }),
            sheets.spreadsheets.values.get({ spreadsheetId, range: 'Don_hang!A1:CF' }),
            sheets.spreadsheets.values.get({ spreadsheetId: spreadsheetHcId, range: 'Data_bom!A1:N' }), // Lấy đến cột N
            sheets.spreadsheets.values.get({ spreadsheetId, range: 'File_BOM_ct!A1:D' })
        ]);

        const data1 = data1Res.data.values || [];
        const data2 = data2Res.data.values || [];
        const data3 = data3Res.data.values || [];
        const data5 = data5Res.data.values || [];

        console.log(`✔️ Đã lấy dữ liệu: ${data1.length} rows (Don_hang_PVC_ct), ${data2.length} rows (Don_hang), ${data3.length} rows (Data_bom), ${data5.length} rows (File_BOM_ct)`);

        const colB = data5.map(row => row[1]).filter(v => v);
        const lastRowWithData = colB.length;
        const d4Value = colB[lastRowWithData - 1];
        if (!d4Value) throw new Error('Không tìm thấy mã đơn hàng trong File_BOM_ct');
        console.log(`✔️ Mã đơn hàng: ${d4Value} (dòng ${lastRowWithData})`);

        const donHang = data2.slice(1).find(row => row[5] === d4Value || row[6] === d4Value);
        if (!donHang) throw new Error(`Không tìm thấy đơn hàng với mã: ${d4Value}`);

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

        hValues.forEach(hObj => {
            const hValue = hObj.hValue;
            const matchingRows = data3.filter(row => row[0] === hValue); // Tìm tất cả row có column A = hValue
            if (matchingRows.length > 0) {
                let isMainRowProcessed = false; // Cờ để chỉ paste columnsToCopyBase vào row chính (cột C = hValue)
                matchingRows.forEach((matchingRow, index) => {
                    let dataFromBN = matchingRow.slice(1, 13); // B:N (index 1 đến 13)
                    let newRow = [...dataFromBN];

                    // Chỉ paste columnsToCopyBase vào row có cột C = hValue (row chính)
                    if (!isMainRowProcessed && newRow[1] === hValue) {
                        const targetValues = columnsToCopyBase.map(i => hObj.rowData[i - 1] || '');
                        newRow.splice(4, 9, ...targetValues); // Ghép vào từ cột E (index 4), thay 9 ô
                        isMainRowProcessed = true;

                        // Tính công thức giống Sheets
                        const rong = parseFloat(newRow[5] || 0); // F: Rộng
                        const cao = parseFloat(newRow[6] || 0); // G: Cao
                        const sl_soi = parseFloat(newRow[7] || 0); // H: SL sợi
                        const sl_bo = parseFloat(newRow[9] || 0); // J: SL bộ

                        newRow[8] = (cao / 1000) * sl_bo; // I: Số lượng
                        newRow[11] = (rong / 1000) * sl_bo; // L: Tổng SL sợi
                        newRow[10] = (rong * cao * sl_soi / 1000000) * newRow[8]; // K: Tổng m2
                        newRow[12] = targetValues[8] || newRow[12] || ''; // M: Ghi chú
                    }

                    tableData.push({
                        stt: hObj.stt,
                        row: newRow
                    });
                    console.log(`✔️ Đã thêm row ${index + 1} cho hValue ${hValue}:`, JSON.stringify(newRow));
                });
            } else {
                console.warn(`⚠️ Không tìm thấy hValue ${hValue} trong Data_bom cột A`);
            }
        });

        console.log('📋 tableData:', JSON.stringify(tableData, null, 2));

        const matchingRows = data2.slice(1).filter(row => row[5] === d4Value || row[6] === d4Value);
        const l4Value = matchingRows[0] ? (matchingRows[0][8] || '') : '';
        const d5Values = matchingRows.flatMap(row => row[83] || []).filter(v => v).join(', ');
        const h5Values = matchingRows.flatMap(row => row[36] || []).filter(v => v).join(', ');
        const h6Values = matchingRows.flatMap(row => row[37] || []).filter(v => v).join(', ');
        const d6Values = matchingRows
            .flatMap(row => row[48] ? new Date(row[48]).toLocaleDateString('vi-VN', { day: '2-digit', month: '2-digit', year: 'numeric' }) : [])
            .filter(v => v)
            .join('<br>');

        const tableDataFrom7 = tableData.slice(0); // Lấy toàn bộ tableData
        console.log('📋 tableDataFrom7:', JSON.stringify(tableDataFrom7, null, 2));

        const uniqueB = [...new Set(tableDataFrom7.map(item => item.row[1]).filter(v => v && v !== 'Mã SP' && v !== 'Mã vật tư sản xuất'))];
        const uniqueC = [...new Set(tableDataFrom7.map(item => item.row[2]).filter(v => v && v !== 'Mã vật tư xuất kèm' && v !== 'Mã vật tư sản xuất'))];
        console.log('📋 uniqueB:', uniqueB);
        console.log('📋 uniqueC:', uniqueC);

        const summaryDataB = uniqueB.map((b, i) => {
            const sum = tableDataFrom7
                .filter(item => item.row[1] === b || item.row[2] === b)
                .reduce((sum, item) => sum + (item.row[8] || item.row[9] || item.row[10] || item.row[11] || 0), 0);
            const desc = tableDataFrom7.find(item => item.row[1] === b || item.row[2] === b)?.row[3] || '';
            return { stt: i + 1, b, sum, desc };
        });
        const summaryDataC = uniqueC.map((c, i) => {
            const sum = tableDataFrom7
                .filter(item => item.row[1] === c || item.row[2] === c)
                .reduce((sum, item) => sum + (item.row[10] || 0), 0);
            const desc = tableDataFrom7.find(item => item.row[1] === c || item.row[2] === c)?.row[3] || '';
            return { stt: summaryDataB.length + i + 1, c, sum, desc };
        });

        console.log(`✔️ Tạo ${summaryDataB.length} mục B và ${summaryDataC.length} mục C trong bảng tổng hợp.`);

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
            hasDataE: tableDataFrom7.some(item => item.row[4]),
            hasDataI: tableDataFrom7.some(item => item.row[8]),
            hasDataJ: tableDataFrom7.some(item => item.row[9]),
            lastRowWithData
        };
    } catch (err) {
        console.error('❌ Lỗi trong prepareYcvtData:', err.stack || err.message);
        throw err;
    }
}

export { prepareYcvtData };