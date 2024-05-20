import exceljs from "exceljs/index";
import saveAs from 'file-saver';
import GC from "@grapecity/spread-sheets";

const workbook = new exceljs.addWorkbook();
workbook.creator = 'User';
workbook.created = new Date();

const sheet = workbook.addWorksheet('First');

document.getElementById("export").onclick = function () {
    var fileName = $("#exportFileName").val();
    var json = JSON.stringify(workbook.toJSON());
    workbook.export(function (blob) {
        // save blob to a file
        saveAs(blob, fileName);
    }, function (e) {
        console.log(e);
    }, {
        fileType: GC.Spread.Sheets.FileType.excel
    });
}
