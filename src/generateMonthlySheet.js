const items = require('./expense-items');
const { workbook, worksheet } = require('./workbook');
const { MONTHS, COLUMNS, WEEKDAY } = require('./constants');

var headerStyle = workbook.createStyle({
  font: {
    bold: true,
  },
  border: {
      bottom: { style: 'thin' },
  }
});

var totalStyle = workbook.createStyle({
    font: {
        bold: true,
    },
    border: {
        top: {
            style: 'thin',
        },
        bottom: {
            style: 'thin',
        },
    }
})


const DATE_COLUMN = 1; 
const DAY_COLUMN = 2;
const TOTAL_COLUMN = DAY_COLUMN + items.length + 1;


const dailyTotalFormula = (rowIndex, columnNames) => items.map(({ expense }, idx) => {
  const columnIdx = idx + DAY_COLUMN;
  const col = columnNames[columnIdx];
  return `${(expense ? '-': '+')}${col}${rowIndex}`
}).join('');



function generateMonthlySheet(monthIndex, yearIndex, rowIndex = 1) {
    const headerRowIndex = rowIndex;

    // Month Name
    worksheet.cell(headerRowIndex, DATE_COLUMN)
        .string(MONTHS[monthIndex])
        .style(headerStyle);

    worksheet.cell(headerRowIndex, DAY_COLUMN)
      .string('Day of Week')
      .style(headerStyle);


    // Item Headers
    items.forEach((item, idx) => {
        worksheet
            .cell(headerRowIndex, idx + DATE_COLUMN + DAY_COLUMN)
            .string(item.name)
            .style(headerStyle);
    });

    // Total Column
    worksheet
      .cell(headerRowIndex, TOTAL_COLUMN)
      .string('Total')
      .style(headerStyle);

    const startDate = Date.UTC(yearIndex, monthIndex, 1);
    const endDate = Date.UTC(yearIndex, monthIndex + 1, 0); // get correct end date

    const start = new Date(startDate) //clone
    const end = new Date(endDate) //clone 

    // Month Date 
    let rowIdx = headerRowIndex;
    while(end >= start) {
      rowIdx += 1;
      worksheet.cell(rowIdx, DATE_COLUMN)
        .date(start);
      worksheet.cell(rowIdx, DAY_COLUMN).string(WEEKDAY[start.getDay()]);

      const f = dailyTotalFormula(rowIdx, COLUMNS);
      worksheet.cell(rowIdx, TOTAL_COLUMN)
        .formula(f)
      
      start.setDate(start.getDate() + 1);
    }

    // Monthly Sub Total
    const totalRowIdx = rowIdx + 1;
    
    worksheet.cell(totalRowIdx, DATE_COLUMN)
        .string('Monthly Total')
        .style(totalStyle);
    worksheet.cell(totalRowIdx, DAY_COLUMN).style(totalStyle);

    // Total
    items.forEach((_, idx) => {
        const columnIdx = idx + DAY_COLUMN; // First column date
        const col = COLUMNS[columnIdx];
        worksheet.cell(totalRowIdx, idx + DATE_COLUMN + DAY_COLUMN)
            .formula(`SUM(${col}${headerRowIndex+1}:${col}${rowIdx})`)
            .style(totalStyle)
    });

    // Total of Total Column
    const totalCol = COLUMNS[TOTAL_COLUMN - 1];
    worksheet.cell(totalRowIdx, TOTAL_COLUMN)
      .formula(`SUM(${totalCol}${headerRowIndex+1}:${totalCol}${rowIdx})`)
      .style(totalStyle);

    return totalRowIdx; // Last Row of month
};



module.exports = generateMonthlySheet;
