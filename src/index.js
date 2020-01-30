
const items = require('./expense-items');


const { workbook, worksheet } = require('./workbook');
const generateMonthlySheet = require('./generateMonthlySheet');


const headerStyle = workbook.createStyle({
  font: {
    bold: true,
  },
  numberFormat: '£#,##0.00; (£#,##0.00); -',
});

worksheet.cell(1, 1)
  .string('December')
  .style(headerStyle);


const headerRowIndex = 1;
items.forEach((item, idx) => {
  worksheet.cell(headerRowIndex, idx + 2)
    .string(item.name)
    .style({ font: { size: 14 }});
});



const startDate = new Date(2019, 11, 1);
const endDate = new Date(2019, 11, 31);


const start = new Date(startDate) //clone
const end = new Date(endDate) //clone 

let rowIdx = headerRowIndex;
while(end >= start) {
  rowIdx += 1;
  worksheet.cell(rowIdx, 1)
    .date(start);
  start.setDate(start.getDate() + 1);
}

const cols = ['B', 'C', 'D']
const totalRowIdx = rowIdx + 1;
items.forEach((item, idx) => {
  const c = cols[idx];
  worksheet.cell(totalRowIdx, idx + 2)
    .formula(`SUM(${c}${headerRowIndex+1}:${c}${rowIdx})`)
});


let startRow = 32; 
[0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11].forEach(m => {
  startRow = generateMonthlySheet(m, 2020, startRow) + 1;
});




 
workbook.write('doc.xlsx');