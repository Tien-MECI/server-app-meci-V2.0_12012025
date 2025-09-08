import { google } from 'googleapis';

async function prepareYcvtData(auth, spreadsheetId, spreadsheetHcId) {
    const sheets = google.sheets({ version: 'v4', auth });

    console.log('▶️ Bắt đầu chuẩn bị dữ liệu cho YCVT...');

    // Lấy dữ liệu từ các sheet
    const [data1Res, data2Res, data3Res, data5Res] = await Promise.all([
        sheets.spreadsheets.values.get({ spreadsheetId, range: 'Don_hang_PVC_ct!A1:AE' }),
        sheets.spreadsheets.values.get({ spreadsheetId, range: 'Don_hang!A1:CF' }),
        sheets.spreadsheets.values.get({ spreadsheetId: spreadsheetHcId, range: 'Data_bom!A1:P' }),
        sheets.spreadsheets.values.get({ spreadsheetId, range: 'File_BOM_ct!A1:D' })
    ]);

    const data1 = data1Res.data.values || [];
    const data2 = data2Res.data.values || [];
    const data3 = data3Res.data.values || [];
    const data5 = data5Res.data.values || [];

    // Tìm d4Value từ cột B của File_BOM_ct
    const colB = data5.map(row => row[1]).filter(v => v);
    const lastRowWithData = colB.length;
    const d4Value = colB[lastRowWithData - 1];
    if (!d4Value) throw new Error('Không tìm thấy mã đơn hàng trong File_BOM_ct');
    console.log(`✔️ Mã đơn hàng: ${d4Value} (dòng ${lastRowWithData})`);

    // Lấy đơn hàng từ Don_hang
    const donHang = data2.slice(1).find(row => row[5] === d4Value || row[6] === d4Value);
    if (!donHang) throw new Error(`Không tìm thấy đơn hàng với mã: ${d4Value}`);

    // Lọc hValues từ Don_hang_PVC_ct
    const hValues = data1.slice(1)
        .filter(row => row[1] === d4Value)
        .map((row, i) => ({
            stt: i + 1,
            hValue: row[7] || '',
            rowData: row
        }));
    console.log(`✔️ Tìm thấy ${hValues.length} sản phẩm với hValue.`);

    // Xử lý bảng chi tiết
    const columnsToCopyBase = [17, 18, 19, 20, 21, 22, 23, 24, 29]; // Cột từ Don_hang_PVC_ct
    let tableData = [];
    let lastProcessedHValue = null;
    let cachedBlock = null;
    hValues.forEach(hObj => {
        const hValue = hObj.hValue;
        if (hValue === lastProcessedHValue && cachedBlock) {
            tableData.push(...cachedBlock.map(row => ({
                stt: hObj.stt,
                row: [...row]
            })));
        } else {
            let z = data3.findIndex(row => row[1] === hValue);
            if (z === -1) return; // Bỏ qua nếu không tìm thấy
            let block = [];
            if (['0S', '0I', 'MD', 'GC', '0N', '0T'].some(str => hValue.includes(str))) {
                let y = data3.slice(z + 1).findIndex(row => row[1] === 'Mã SP') + z + 1;
                if (y < z) return; // Bỏ qua nếu không tìm thấy y
                block = data3.slice(z, y + 1);
            } else {
                let x = data3.slice(0, z + 1).reverse().findIndex(row => row[1] === 'Mã SP');
                x = z - x;
                if (x === -1) return; // Bỏ qua nếu không tìm thấy x
                block = [data3[x], data3[z + 1]].filter(row => row);
            }
            tableData.push(...block.map(row => ({
                stt: hObj.stt,
                row: [...row]
            })));
            cachedBlock = block;
            lastProcessedHValue = hValue;
        }
        // Cập nhật cột E-M
        const targetValues = columnsToCopyBase.map(i => hObj.rowData[i - 1] || '');
        tableData[tableData.length - 1].row.splice(4, 9, ...targetValues);
    });

    // Xử lý metadata từ Don_hang
    const matchingRows = data2.slice(1).filter(row => row[5] === d4Value || row[6] === d4Value);
    const l4Value = matchingRows[0] ? (matchingRows[0][8] || '') : '';
    const d5Values = matchingRows.flatMap(row => row[83] || []).filter(v => v).join(', ');
    const h5Values = matchingRows.flatMap(row => row[36] || []).filter(v => v).join(', ');
    const h6Values = matchingRows.flatMap(row => row[37] || []).filter(v => v).join(', ');
    const d6Values = matchingRows
        .flatMap(row => row[48] ? new Date(row[48]).toLocaleDateString('vi-VN', { day: '2-digit', month: '2-digit', year: 'numeric' }) : [])
        .filter(v => v)
        .join('<br>');

    // Tạo bảng tổng hợp
    const tableDataFrom7 = tableData.slice(6); // Từ row 7
    const uniqueB = [...new Set(tableDataFrom7.map(item => item.row[1]).filter(v => v && v !== 'Mã SP' && v !== 'Mã vật tư sản xuất'))];
    const uniqueC = [...new Set(tableDataFrom7.map(item => item.row[2]).filter(v => v && v !== 'Mã vật tư xuất kèm' && v !== 'Mã vật tư sản xuất'))];

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
}

export { prepareYcvtData };