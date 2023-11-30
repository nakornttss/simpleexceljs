const ExcelJS = require('exceljs');

async function createExcelFile() {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('My Sheet');

    // Define columns
    worksheet.columns = [
        { header: 'Id', key: 'id', width: 10 },
        { header: 'Name', key: 'name', width: 32 },
        { header: 'D.O.B.', key: 'dob', width: 15, style: { numFmt: 'mm/dd/yyyy' } }
    ];

    // Add a row by sparse array (assign to columns A, B & C)
    worksheet.addRow({id: 1, name: 'John Doe', dob: new Date(1970, 1, 1)});

    // Style the first row
    const firstRow = worksheet.getRow(1);
    firstRow.font = { bold: true };
    firstRow.getCell(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFF00' },
    };

    // Write to a file
    await workbook.xlsx.writeFile('ExcelFile.xlsx');
    console.log('Excel file created successfully!');
}

console.log('test');

createExcelFile().catch(err => console.error(err));
