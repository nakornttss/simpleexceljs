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

    // Add a row
    worksheet.addRow({id: 1, name: 'John Doe', dob: new Date(1970, 1, 1)});

    // Style the first row
    const firstRow = worksheet.getRow(1);
    firstRow.font = { bold: true };

    // Apply styles to all cells in the first row
    firstRow.eachCell(cell => {
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFF00' }
        };

        cell.border = {
            top: { style: 'thin', color: { argb: 'FF0000' } },
            left: { style: 'thin', color: { argb: 'FF0000' } },
            bottom: { style: 'thin', color: { argb: 'FF0000' } },
            right: { style: 'thin', color: { argb: 'FF0000' } }
        };
    });

    // Write to a file
    await workbook.xlsx.writeFile('ExcelFileWithStyles.xlsx');
    console.log('Excel file with styles created successfully!');
}

createExcelFile().catch(err => console.error(err));
