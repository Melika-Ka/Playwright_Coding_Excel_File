import exceljs from "exceljs";
// way2
async function excelTest(searchText, replaceText, change, filePath) {
  const newWorkbook = new exceljs.Workbook();
  await newWorkbook.xlsx.readFile(filePath);
  const workSheet = newWorkbook.getWorksheet("Sheet1");
  const output = await readExcel(workSheet, searchText);
  const cell = workSheet.getCell(
    output.row + change.rowChange,
    output.col + change.colChange
  );
  cell.value = replaceText;
  await newWorkbook.xlsx.writeFile("updatedFile.xlsx");
  console.log(output);
  console.log(searchText);
  console.log(cell.value);
}
async function readExcel(workSheet, searchText) {
  let output = { row: -1, col: -1 };

  workSheet.eachRow((row, rowNumber) => {
    row.eachCell((cell, colNumber) => {
      if (cell.value == searchText) {
        output.row = rowNumber;
        output.col = colNumber;
      }
    });
  });
  return output;
}
excelTest("Mango", 350, { rowChange: 0, colChange: 2 }, "download.xlsx");
