const ExcelJs = require('exceljs');

async function excelTest() {
    let output = {row:-1,column:-1};
    const workbook = new ExcelJs.Workbook();
    await workbook.xlsx.readFile("sample.xlsx")
    const workSheet = workbook.getWorksheet('Sheet1');
    workSheet.eachRow((row, rowNumber) => {
        row.eachCell((cell, colNumber) => {
            if(cell.value === "Banana") {
                //console.log(rowNumber,colNumber)
                output.row = rowNumber;
                output.column = colNumber;
            }
        })
    })

    const cell = workSheet.getCell(output.row,output.column);
    cell.value = "Republic";
    await workbook.xlsx.writeFile("sample.xlsx");
}

excelTest();
