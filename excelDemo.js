const ExcelJs = require('exceljs');

const workbook = new ExcelJs.Workbook();
workbook.xlsx.readFile("sample.xlsx").then(function () {
    const workSheet = workbook.getWorksheet('Sheet1');
    workSheet.eachRow((row, rowNumber) => {
        row.eachCell((cell, colNumber) => {
            console.log(cell.value);
        })
    })
});
