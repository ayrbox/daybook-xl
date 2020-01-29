const xl = require('excel4node');

const workbook = new xl.Workbook({
  defaultFont: {
    size: 12,
    name: 'Arial',
    color: '#505050',
  }
});

const worksheet = workbook.addWorksheet('Expense Sheet');

module.exports = {
    workbook,
    worksheet,
};