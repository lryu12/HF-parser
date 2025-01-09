const XLSX = require('xlsx');
const HyperFormula = require('hyperformula').HyperFormula;

async function parseSpreadsheet(filePath) {
  const workbook = XLSX.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  const hfInstance = HyperFormula.buildFromArray(jsonData);

  const parsedData = hfInstance.getSheetValues(0).map((row, rowIndex) => {
    return row.map((value, colIndex) => {
      const formula = hfInstance.getCellFormula({ sheet: 0, row: rowIndex, col: colIndex });
      return formula ? { formula, value } : value;
    });
  });

  return parsedData;
}

module.exports = parseSpreadsheet;
