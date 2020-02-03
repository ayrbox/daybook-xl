const items = require('./expense-items');
const { workbook, worksheet } = require('./workbook');
const { COLUMNS } = require('./constants');

var grandTotalStyle = workbook.createStyle({
    font: {
        bold: true,
        size: 18,
    },
    border: {
        top: {
            style: 'double',
        },
    }
});

function grandTotal(rowIndexes, startAt) {
    worksheet.cell(startAt, 1)
        .string('Grand Total')
        .style(grandTotalStyle);
    worksheet.cell(startAt, 2)
        .style(grandTotalStyle);

    items.forEach((_, idx) => {
        const columnIdx = idx + 2;
        const col = COLUMNS[columnIdx];

        const f = rowIndexes.map((rowIdx) => `${col}${rowIdx}`);

        worksheet.cell(startAt, idx + 3)
            .formula(f.join('+'))
            .style(grandTotalStyle);
    });

    const TOTAL_COLUMN = 1 + items.length + 1;


    // Grand Total of Total Column
    const totalCol = COLUMNS[TOTAL_COLUMN - 1];
    const f = rowIndexes.map((rowIdx) => `${totalCol}${rowIdx}`);
    worksheet.cell(startAt, TOTAL_COLUMN)
      .formula(f.join('+'))
      .style(grandTotalStyle);
} 

module.exports = grandTotal;
