const XLSX = require('xlsx');

function extractColumnBasedFormulas(filePath, targetRow) {
  const workbook = XLSX.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  const formulasForRow = {};

  for (const cellAddress in sheet) {
    if (cellAddress[0] === '!') continue;

    const cell = sheet[cellAddress];
    const rowNumber = parseInt(cellAddress.replace(/^\D+/, ''), 10);

    if (rowNumber === targetRow && cell.f) {
      const columnLetter = cellAddress.replace(/\d+$/, '');

      const adjustedFormula = cell.f.replace(/[A-Z]+\d+/g, match => {
        return match.replace(/\d+$/, '');
      });

      formulasForRow[columnLetter] = adjustedFormula;
    }
  }

  return formulasForRow;
}


async function parseSpreadsheet(filePath) {
  const workbook = XLSX.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  const sheetArray = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  const formulaArray = XLSX.utils.sheet_to_formulae(sheet);
  console.log(formulaArray);

  const headers = sheetArray[3];
  const formulas = extractColumnBasedFormulas(filePath, 5);

  console.log('Headers:', headers);
  console.log('Formulas in first row:', formulas);

  return { headers, formulas };
}

module.exports = parseSpreadsheet;
