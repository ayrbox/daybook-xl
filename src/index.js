const { EMPTY_LINES } = require('./constants');
const { workbook } = require('./workbook');
const generateMonthlySheet = require('./generateMonthlySheet');
const grandTotal = require('./grandTotal');

const totalRowIndexes = [];

const fiscalMonths = [
  { month: 3, year: 2020 },
  { month: 4, year: 2020 },
  { month: 5, year: 2020 },
  { month: 6, year: 2020 },
  { month: 7, year: 2020 },
  { month: 8, year: 2020 },
  { month: 9, year: 2020 },
  { month: 10, year: 2020 },
  { month: 11, year: 2020 },
  { month: 0, year: 2021 },
  { month: 1, year: 2021 },
  { month: 2, year: 2021 },
];

let startRow = 1; 
fiscalMonths.forEach(({ month, year }) => {
  const totalRowIndex = generateMonthlySheet(month, year, startRow);
  totalRowIndexes.push(totalRowIndex);
  startRow = totalRowIndex + EMPTY_LINES + 1;
});



const grandTotalRowIdx = startRow;

grandTotal(totalRowIndexes, grandTotalRowIdx);

workbook.write('doc.xlsx');