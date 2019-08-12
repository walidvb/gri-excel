
const { computeStepPrice } =  require('./pricing');
const data = require('./data')

const COLUMNS = [
 {header: 'No', key: 'no', width: 10},
 {header: 'Descriptif', key: 'description', width: 40},
 {header: 'Quantité', key: 'quantity', width: 10},
 {header: 'Mesure', key: 'unit', width: 10},
 {header: 'Prix Unitaire', key: 'unit_price', width: 20},
 {header: 'Total', key: 'total', width: 20}
]
const colToLetter = (n) => 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'[n-1]
const lastCol = colToLetter(COLUMNS.length)
class Excelor{
  constructor(data){
    this.data = data
    this.rooms = data.version.rooms
    this.date = data.version.created_at
  }
  async createDocument(){
    this.initWorkBook()
    this.sheet = this.workbook.addWorksheet('My Sheet')
    this.sheet.columns = COLUMNS
    this.addInfo()
    const filename = "now.xls"
    await this.workbook.xlsx.writeFile(filename)
    return filename
  }
  addInfo(){
    // addHeader(sheet)
    this.addColumnNames()
    this.rooms.forEach(this.addRoom.bind(this))
    // addTotal(sheet)
  }
  addColumnNames(){
    this.sheet.addRow(COLUMNS);
  }
  addRoom({ name, steps }){
    addRoomTitle.call(this)
    const firstRoomRow = this.sheet.lastRow._number + 1
    steps.forEach(addStep.bind(this))
    const lastRoomRow = this.sheet.lastRow._number
    addTotal.call(this)

    function addTotal(){
      this.sheet.addRow(['','','','','Total', ''])
      const totalCol = colToLetter(this.sheet.getColumn('total')._number)
      const formula = `SUM(${totalCol}${firstRoomRow}:${totalCol}${lastRoomRow})`
      const row = this.sheet.lastRow
      row.getCell('total').value = { formula };
    }
    function addRoomTitle(){
      this.sheet.addRow(['', name])
      const row = this.sheet.lastRow
      const number = row._number
      row.font = { bold: true }
      this.sheet.mergeCells(`B${number}:${lastCol}${number}`)
    }
    function addStep(step) {
      const { id, quantity, unit_price, description, unit, price } = step
      const stepRow = [
        id,
        description,
        quantity,
        unit,
        unit_price,
        ''
      ]
      this.sheet.addRow(stepRow)
      const row = this.sheet.lastRow
      const priceFormula = `${row.getCell('quantity')._address} * ${row.getCell('unit_price')._address}`
      const totalFormula = `MAX(${priceFormula}, ${price})`
      row.getCell('total').value = { formula: totalFormula }
    }
  }
  initWorkBook(){
    const Excel = require('exceljs/modern.nodejs');
    var workbook = new Excel.Workbook();
    workbook.creator = 'Me';
    workbook.lastModifiedBy = 'Her';
    workbook.created = new Date(1985, 8, 30);
    workbook.modified = new Date();
    workbook.lastPrinted = new Date(2016, 9, 27);
    this.workbook = workbook
  }
}

new Excelor(data).createDocument()
module.exports = Excelor