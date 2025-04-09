const ExcelJs = require('exceljs');


async function writeExcelTest(searchText,replaceText,change,filePath) {
    
    const workbook = new ExcelJs.Workbook();
    await workbook.xlsx.readFile(filePath)
    const workSheet = workbook.getWorksheet('Sheet1');
    const output = await readExcel(workSheet,searchText);
    if(output.row === -1 && output.column === -1) {
        console.log("Search text not found");
        return;
    }
    const cell = workSheet.getCell(output.row,output.column+change.colChange);
    cell.value = replaceText;
    await workbook.xlsx.writeFile(filePath);
}

async function readExcel(workSheet,searchText){
    let output = {row:-1,column:-1};
    workSheet.eachRow((row, rowNumber) => {
        row.eachCell((cell, colNumber) => {
            if(cell.value === searchText) {
                output.row = rowNumber;
                output.column = colNumber;
            }
        })
    })
    //console.log(output);
    return output;
}

//change price to 350
writeExcelTest("Mango","350",{rowChange:0,colChange:2},"sample.xlsx");
