const items = require('./expense-items');
const { worksheet } = require('./workbook');
const { COLUMNS } = require('./constants');


function grandTotal(rowIndexes, startAt) {
    worksheet.cell(startAt, 1)
        .string('Grand Total');

    items.forEach((_, idx) => {
        const columnIdx = idx + 1;
        const col = COLUMNS[columnIdx];

        const f = rowIndexes.map((rowIdx) => {
            return `${col}${rowIdx}`;
        });

        worksheet.cell(startAt, idx + 2)
            .formula(f.join('+'));
    });
} 

module.exports = grandTotal;