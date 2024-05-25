import * as FileSaver from './scripts/file-saver/src/FileSaver.js';
import ExcelJS from './scripts/exceljs/excel.js';

console.log('1');

const workbook = new ExcelJS.Workbook();
workbook.creator = 'User';
workbook.created = new Date();
workbook.modified = new Date();

const flightPlan = workbook.addWorksheet('flightPlan', {
    properties: { tabColor: { argb:'00B050'}}, views: [{state: 'frozen', ySplit:1}]
});

flightPlan.columns = [
    {header: 'MSA', key: 'msa', width: 10, style: {font: {name: 'Cambria', bold: true, underline: true, size: 12 }}},
    {header: 'ALT', key: 'alt', width: 10, style: {font: {name: 'Cambria', bold: true, underline: true, size: 12 }}},
    {header: 'TAS', key: 'tas', width: 10, style: {font: {name: 'Cambria', bold: true, underline: true, size: 12 }}},
]

console.log('2');

document.getElementById('saveFile').addEventListener('click', async function(e) {
    try {
        const buffer = await workbook.xlsx.writeBuffer();
        const fileType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
        let EXCEL_EXTENSION = '.xlsx';
        const blob = new Blob([buffer], {type: fileType});

        FileSaver.saveAs(blob, `Excel_Spreadsheet` + EXCEL_EXTENSION);
    } catch (err) {
        console.error(err);
    }

console.log('3');
});

console.log('4');
