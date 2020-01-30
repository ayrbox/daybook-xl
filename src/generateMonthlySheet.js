const items = require('./expense-items');
const { workbook, worksheet } = require('./workbook');
const { MONTHS, COLUMNS } = require('./constants');

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
const TOTAL_COLUMN = DATE_COLUMN + items.length + 1;


const dailyTotalFormula = (rowIndex, columnNames) => items.map(({ expense }, idx) => {
  const columnIdx = idx + 1;
  const col = columnNames[columnIdx];
  return `${(expense ? '-': '+')}${col}${rowIndex}`
}).join('');



function generateMonthlySheet(monthIndex, yearIndex, rowIndex = 1) {
    const headerRowIndex = rowIndex;

    // Month Name
    worksheet.cell(headerRowIndex, 1)
        .string(MONTHS[monthIndex])
        .style(headerStyle);

    // Item Headers
    items.forEach((item, idx) => {
        worksheet
            .cell(headerRowIndex, idx + 2)
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
      worksheet.cell(rowIdx, 1)
        .date(start)

      const f = dailyTotalFormula(rowIdx, COLUMNS);
      worksheet.cell(rowIdx, TOTAL_COLUMN)
        .formula(f)

      
      start.setDate(start.getDate() + 1);
    }

    // Monthly Sub Total
    const totalRowIdx = rowIdx + 1;
    
    worksheet.cell(totalRowIdx, 1)
        .string('Monthly Total')
        .style(totalStyle);


    // Total
    items.forEach((_, idx) => {
        const columnIdx = idx + 1; // First column date
        const col = COLUMNS[columnIdx];
        worksheet.cell(totalRowIdx, idx + 2)
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
