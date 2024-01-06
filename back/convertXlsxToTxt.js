const XLSX = require('xlsx');

const ExcelJS = require('exceljs');
const fs = require('fs');

//ChatGPT response
function convertXlsxToTxt(filePath) {

    try {
        const workbook = XLSX.readFile(filePath);
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        let lastLine = false;

        const range = XLSX.utils.decode_range(sheet['!ref']);

        let txtData = '';
        let lineTxt = '';

        txtData += '\n';

        for (let row = range.s.r; row <= range.e.r && !lastLine; row++) {
            lineTxt = '';
            if (row !== 7) {
                for (let col = range.s.c; col <= range.e.c; col++) {
                    const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                    const cellValue = sheet[cellAddress] ? sheet[cellAddress].v : ' ';
                    switch (col) {
                        case 0:
                            if (row === 1 && cellValue === 'AN') {
                                lineTxt += '' + cellValue;
                            }
                            else if (row === 2) {
                                lineTxt += ExcelDateToJSDate(cellValue, true);
                            }
                            else if (row > 7) {
                                if (row === 8) {
                                    lineTxt += '         ' + cellValue + '    ';
                                }
                                else {
                                    const nextCellAddress = XLSX.utils.encode_cell({ r: row, c: col + 1 });
                                    const nextCellValue = sheet[nextCellAddress] ? sheet[nextCellAddress].v : ' ';

                                    if (typeof cellValue === 'number' && nextCellValue !== 'asientos') {
                                        const date = ExcelDateToJSDate(cellValue);
                                        txtData += '       ' + date + ' ';
                                    }
                                    else if (typeof cellValue === 'number' && nextCellValue === 'asientos') {
                                        lastLine = true;
                                        lineTxt += '          ' + cellValue + ' ';
                                    }
                                    else {
                                        lineTxt += '                  ';
                                    }
                                }
                            }
                            else {
                                lineTxt += cellValue + ' ';
                            }
                            break;
                        case 1:
                            if (row === 8) {
                                lineTxt += cellValue + '                              '
                            }
                            else if (row !== 2 && row !== 4) {
                                if (cellValue === 'asientos') {
                                    lineTxt += cellValue.padEnd(21, ' ');
                                }
                                else {
                                    lineTxt += cellValue.padEnd(16, ' ');
                                }
                            }
                            break;
                        case 2, 3, 4, 8:
                            break;
                        case 5:
                            if (row === 2) {
                                lineTxt += '     ' + cellValue;
                            }
                            break;
                        case 6:
                            if (row === 8) {
                                lineTxt += cellValue + '               '
                            }
                            else if (row === 4) {
                                lineTxt += '                     ' + cellValue;
                            }
                            else if (row > 8) {
                                let value = parseFloat(cellValue.toString()).toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,');
                                lineTxt += value.toString() !== 'NaN' ? value.toString().padStart(27, ' ') : '                        ';
                            }
                            break;
                        case 7:
                            if (row === 8) {
                                lineTxt += cellValue + '      '
                            }
                            else if (row > 8) {
                                let value = parseFloat(cellValue.toString()).toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,');
                                if (lastLine) {
                                    lineTxt += value.toString() !== 'NaN' ? value.toString().padStart(20, ' ') : '                    ';
                                }
                                else {
                                    lineTxt += value.toString() !== 'NaN' ? value.toString().padStart(23, ' ') : '                    ';
                                }
                            }
                            break;
                        case 9:
                            if (row === 8) {
                                lineTxt += cellValue + '    '
                            }
                            else if (row > 8) {
                                lineTxt += cellValue.toString() !== ' ' ? cellValue.toString().padEnd(13, ' ') : '             ';
                            }
                            break;
                        case 11:
                            if (row === 8) {
                                lineTxt += cellValue
                            }
                            else if (row > 8) {
                                lineTxt += cellValue;
                            }
                            break;
                    }
                }
                txtData += lineTxt + '\n';

                if (lastLine) {
                    txtData += '\n' + lineTxt + '\n\n'
                }
                // Add a line break after the 8th row, the headers row
                if (row === 8 || row === 4) {
                    txtData += '\n';
                }
            }
        }

        const txtFilePath = filePath.replace('.xlsx', '_output.txt');
        fs.writeFileSync(txtFilePath, txtData);

        return { convertion: true, error: null };
    }
    catch (error) {
        return { convertion: false, error: error }
    }
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