const xl = require('excel4node');


const items = require('./expense-items');





const wb = new xl.Workbook();
const ws = wb.addWorksheet('Sheet 1');


var style = wb.createStyle({
  font: {
    color: '#FF0800',
    size: 12,
  },
  numberFormat: '£#,##0.00; (£#,##0.00); -',
});

ws.cell(1, 1)
  .string('December')
  .style(style);


const headerRowIndex = 1;
items.forEach((item, idx) => {
  ws.cell(headerRowIndex, idx + 2)
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
  ws.cell(rowIdx, 1)
    .date(start);

  start.setDate(start.getDate() + 1);
}

const cols = ['B', 'C', 'D']
const totalRowIdx = rowIdx + 1;
items.forEach((item, idx) => {
  const c = cols[idx];
  ws.cell(totalRowIdx, idx + 2)
    .formula(`SUM(${c}${headerRowIndex+1}:${c}${rowIdx})`)
});






// ws.cell(1, 1)
//   .number(100)
//   .style(style);
 
// // Set value of cell B1 to 200 as a number type styled with paramaters of style
// ws.cell(1, 2)
//   .number(200)
//   .style(style);
 
// // Set value of cell C1 to a formula styled with paramaters of style
// ws.cell(1, 3)
//   .formula('SUM(A1:B1)')
//   .style(style);
 
// // Set value of cell A2 to 'string' styled with paramaters of style
// ws.cell(2, 1)
//   .string('string')
//   .style(style);
 
// // Set value of cell A3 to true as a boolean type styled with paramaters of style but with an adjustment to the font size.
// ws.cell(3, 1)
//   .bool(true)
//   .style(style)
//   .style({font: {size: 14}});
 
wb.write('doc.xlsx');