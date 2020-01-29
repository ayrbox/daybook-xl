const items = require('./expense-items');
const { workbook, worksheet } = require('./workbook');
const { MONTHS, COLUMNS } = require('./constants');

var headerStyle = workbook.createStyle({
  font: {
    bold: true,
  },
});

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

    const startDate = new Date(yearIndex, monthIndex, 1);
    const endDate = new Date(yearIndex, monthIndex + 1, 0); // get correct end date

    const start = new Date(startDate) //clone
    const end = new Date(endDate) //clone 

    // Month Date 
    let rowIdx = headerRowIndex;
    while(end >= start) {
        rowIdx += 1;
        worksheet.cell(rowIdx, 1)
            .date(start);
        start.setDate(start.getDate() + 1);
    }

    // Monthly Sub Total
    const totalRowIdx = rowIdx + 1;
    items.forEach((_, idx) => {
        const columnIdx = idx + 1; // First column date
        const col = COLUMNS[columnIdx];
        worksheet.cell(totalRowIdx, idx + 2)
            .formula(`SUM(${col}${headerRowIndex+1}:${col}${rowIdx})`)
    });

    return totalRowIdx; // Last Row of month
};



module.exports = generateMonthlySheet;