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
async function excelTest() {
  const newWorkbook = new exceljs.Workbook();
  await newWorkbook.xlsx.readFile("download.xlsx");
  const workSheet = newWorkbook.getWorksheet("Sheet1");
  await workSheet.eachRow((row, rowNumber) => {
    row.eachCell((cell, colNumber) => {
      console.log(cell.value);
    });
  });
}
excelTest();
