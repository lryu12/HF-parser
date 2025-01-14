const XLSX = require('xlsx');

// Extract relative formulas from each colunn
function extractColumnBasedFormulas(filePath, targetRow) {
  const workbook = XLSX.readFile(filePath);
  // Get name of the first worksheet
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  const formulasForRow = {};

  // Loops over each cell of the sheet (ex. A1, A2, B1, B2 ...)
  for (const cellAddress in sheet) {
    if (cellAddress[0] === '!') continue; // Skip metadata

    const cell = sheet[cellAddress];
    const rowNumber = parseInt(cellAddress.replace(/^\D+/, ''), 10); // Extract row number

    // Check if the cell belongs to the target row (currently only works with row 5 but can 
    // use multiple rows if needed)
    if (rowNumber === targetRow && cell.f) {
      // Extract the column letter (e.g., A5 -> A)
      const columnLetter = cellAddress.replace(/\d+$/, '');

      // Convert the formula to column-based references
      const adjustedFormula = cell.f.replace(/[A-Z]+\d+/g, match => {
        return match.replace(/\d+$/, ''); // Strip row number, keep column letter
      });

      // Store the formula for this column
      formulasForRow[columnLetter] = adjustedFormula;
    }
  }

  return formulasForRow;
}


async function parseSpreadsheet(filePath) {
  // Load the workbook and select the first sheet
  const workbook = XLSX.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  // Convert the sheet to a 2D array
  const sheetArray = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  const formulaArray = XLSX.utils.sheet_to_formulae(sheet);
  console.log(formulaArray);

  // Extract headers & assume first row contains headers
  // the example file contains 3 empty rows
  const headers = sheetArray[3];

  // Extract formulas from the first row (5 for the example file)
  const formulas = extractColumnBasedFormulas(filePath, 5);

  console.log('Headers:', headers);
  console.log('Formulas in first row:', formulas);

  return { headers, formulas };
}

module.exports = parseSpreadsheet;
