import exceljs from "exceljs";
// way1

// const newWorkbook = new exceljs.Workbook();
// newWorkbook.xlsx.readFile("download.xlsx").then(() => {
//   const workSheet = newWorkbook.getWorksheet("Sheet1");
//   workSheet.eachRow((row, rowNumber) => {
//     row.eachCell((cell, colNumber) => {
//       console.log(cell.value);
//     });
//   });
// });

// way2
let output = { row: -1, col: -1 };
async function excelTest() {
  const newWorkbook = new exceljs.Workbook();
  await newWorkbook.xlsx.readFile("download.xlsx");
  const workSheet = newWorkbook.getWorksheet("Sheet1");
  workSheet.eachRow((row, rowNumber) => {
    row.eachCell((cell, colNumber) => {
      if (cell.value == "Apple") {
        // console.log(rowNumber);
        // console.log(colNumber);
        output.row = rowNumber;
        output.col = colNumber;
      }
    });
  });
  const cell = workSheet.getCell(output.row, output.col);
  cell.value = "Iphone";
  await newWorkbook.xlsx.writeFile("updatedFile.xlsx");
}
excelTest();
