const path = require('path');
const xlsx = require('xlsx');

function readExcelData(excelPath) {
  const fullPath = path.join(__dirname, excelPath);
  const workbook = xlsx.readFile(fullPath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  return xlsx.utils.sheet_to_json(worksheet);
}

module.exports = { readExcelData };
