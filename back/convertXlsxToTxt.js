const XLSX = require('xlsx');

const ExcelJS = require('exceljs');
const fs = require('fs');

/* async function convertXlsxToTxt(filePath) {
    const workbook = new ExcelJS.Workbook();

    // Load the existing Excel file
    await workbook.xlsx.readFile(filePath);

    // Get the desired sheet
    const sheet = workbook.getWorksheet('Report');

    // Create a string to store the TXT content
    let txtContent = '';

    // Iterate through rows in the sheet
    sheet.eachRow((row, rowNumber) => {
        // Customize the logic based on your VBA code
        if (row.getCell(2).value === 'asientos') {
            txtContent += `${row.getCell(1).text} ${row.getCell(2).text}\n`;
        } else if (row.getCell(2).value === 'Fecha') {
            // Customize for 'Fecha' row
            txtContent += ` ${row.getCell(2).text}   ${row.getCell(5).text} ${row.getCell(10).text} ${row.getCell(13).text} ${row.getCell(16).text} ${row.getCell(18).text}\n`;
        } else if (row.getCell(2).value !== 'Fecha' && row.getCell(2).value !== 'asientos') {
            // Customize for other rows
            txtContent += ` ${row.getCell(2).text} ${row.getCell(5).text} ${row.getCell(10).text} ${row.getCell(13).text} ${row.getCell(16).text} ${row.getCell(18).text}\n`;
        }
    });

    // Save the TXT content to a file

    const txtFilePath = filePath.replace('.xlsx', '_output.txt');
    fs.writeFileSync(txtFilePath, txtContent, 'utf-8');

    console.log('Conversion completed. Check output.txt');
} */


//ChatGPT response
function convertXlsxToTxt(filePath) {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    const range = XLSX.utils.decode_range(sheet['!ref']);

    let txtData = '';

    txtData += '\n';

    for (let row = range.s.r; row <= range.e.r; row++) {
        if (row !== 7) {
            for (let col = range.s.c; col <= range.e.c; col++) {
                const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                const cellValue = sheet[cellAddress] ? sheet[cellAddress].v : ' ';

                switch (col) {
                    case 0:
                        if (row === 2) {
                            txtData += ExcelDateToJSDate(cellValue, true);
                        }
                        else if (row > 7) {
                            if (row === 8) {
                                txtData += '         ' + cellValue + '    ';
                            }
                            else {
                                const nextCellAddress = XLSX.utils.encode_cell({ r: row, c: col + 1 });
                                const nextCellValue = sheet[nextCellAddress] ? sheet[nextCellAddress].v : ' ';

                                if (typeof cellValue === 'number' && nextCellValue !== 'asientos') {
                                    const date = ExcelDateToJSDate(cellValue);
                                    txtData += '       ' + date + ' ';
                                }
                                else {
                                    txtData += '                ' + cellValue + ' ';
                                }
                            }
                        }
                        else {
                            txtData += cellValue + ' ';
                        }
                        break;
                    case 1:
                        if (row === 8) {
                            txtData += cellValue + '                              '
                        }
                        else if (row !== 2 && row !== 4) {
                            txtData += cellValue.padEnd(16, ' ');
                        }
                        break;
                    case 2, 3, 4, 8:
                        break;
                    case 5:
                        if (row === 2) {
                            txtData += '     ' + cellValue;
                        }
                        break;
                    case 6:
                        if (row === 8) {
                            txtData += cellValue + '               '
                        }
                        else if (row === 4) {
                            txtData += '                     ' + cellValue;
                        }
                        else if (row > 8) {
                            let value = parseFloat(cellValue.toString()).toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,');
                            txtData += value.toString() !== 'NaN' ? value.toString().padStart(27, ' ') : '                        ';
                        }
                        break;
                    case 7:
                        if (row === 8) {
                            txtData += cellValue + '      '
                        }
                        else if (row > 8) {
                            let value = parseFloat(cellValue.toString()).toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,');
                            txtData += value.toString() !== 'NaN' ? value.toString().padStart(23, ' ') : '                    ';
                        }
                        break;
                }
            }
            txtData += '\n';
            // Add a line break after the 8th row, the headers row
            if (row === 8 || row === 4) {
                txtData += '\n';
            }
        }
    }

    const txtFilePath = filePath.replace('.xlsx', '_output.txt');
    fs.writeFileSync(txtFilePath, txtData);

    console.log('Conversion complete. TXT file saved at:', txtFilePath);
}

function ExcelDateToJSDate(date, dateTime = false) {
    return formatDate(new Date(Math.round((date - 25569) * 86400 * 1000)), dateTime);
}

function formatDate(date, dateTime) {
    const months = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']
    let string = `${date.getUTCDate() > 9 ? date.getUTCDate() : '0' + date.getUTCDate()}/${months[date.getUTCMonth()]}/${date.getUTCFullYear()}`
    if (dateTime) {
        string += ` (${date.getUTCHours() > 9 ? date.getUTCHours() : '0' + date.getUTCHours()}:${date.getUTCMinutes() > 9 ? date.getUTCMinutes() : '0' + date.getUTCMinutes()}:${date.getUTCSeconds() > 9 ? date.getUTCSeconds() : '0' + date.getUTCSeconds()})`
    }
    return string
}

module.exports = convertXlsxToTxt;