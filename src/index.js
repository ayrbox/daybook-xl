const { EMPTY_LINES } = require('./constants');
const { workbook } = require('./workbook');
const generateMonthlySheet = require('./generateMonthlySheet');
const grandTotal = require('./grandTotal');

const totalRowIndexes = [];

let startRow = 1; 
[0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11].forEach(m => {
  const totalRowIndex = generateMonthlySheet(m, 2020, startRow);
  totalRowIndexes.push(totalRowIndex);
  startRow = totalRowIndex + EMPTY_LINES + 1;
});


const grandTotalRowIdx = startRow;

grandTotal(totalRowIndexes, grandTotalRowIdx);

workbook.write('doc.xlsx');