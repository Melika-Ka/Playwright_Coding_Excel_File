import exceljs from "exceljs";
// way1

const newWorkbook = new exceljs.Workbook();
newWorkbook.xlsx.readFile("download.xlsx").then(() => {
  const workSheet = newWorkbook.getWorksheet("Sheet1");
  workSheet.eachRow((row, rowNumber) => {
    row.eachCell((cell, colNumber) => {
      console.log(cell.value);
    });
  });
});
