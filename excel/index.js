const Excel = require('exceljs');
const fs = require('fs');
const maybeAddDiscount = require('./addDiscounts')

const COLUMNS = [
 {header: 'No', key: 'no', width: 5},
 {header: 'Descriptif', key: 'description', width: 70},
 {header: 'Quantité', key: 'quantity', width: 10},
 {header: 'Mesure', key: 'unit', width: 10},
 {header: 'Prix Unitaire', key: 'price_display', width: 10},
 {header: 'Total', key: 'total', width: 10},
 {header: '', key: 'price', width: 10},
]
const colToLetter = (n) => 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'[n-1]
const letterToNumber = (l) => 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.indexOf(l)
const LAST_COL = colToLetter(COLUMNS.length - 1)

const borderTop = { border: { top: { style: 'thin', color: { argb: '#DDDDDDDD' } } } }
const border = (color = '#DDDDDDDD') =>  ({ border: ['top', 'bottom', 'left', 'right'].reduce((prev, curr) => ({ ...prev, [curr]: { style: 'thin', color: { argb: color } } }), {}) })
const wrap = { alignment: { wrapText: true } }
const fill = (color) => ({
    fill: {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: color },
    }
})
const formatCells = (format, row, range) => {
  const number = row._number
  const from = letterToNumber(range[0])
  const to = letterToNumber(range[1])
  for (let i = from; i <= to; i++){
    const cell = row.getCell(colToLetter(i+1))
    for(let f in format){
      cell[f] = format[f]
    }
  }
}
class Excelor{
  constructor(data){
    const { version, user, ...project } = data
    this.project = project
    this.user = user
    this.version = version
    this.rooms = version.rooms
    this.date = version.created_at
    this.cellsThatAreTotal = []
    this.currentStepIndex = 1

    this.maybeAddDiscount = maybeAddDiscount.bind(this)
  }
  async createDocument(){
    const { id: pID } = this.project
    const { vID } = this.version
    this.id = `${pID}-${vID}`
    this.initWorkBook()
    this.sheet = this.workbook.addWorksheet('My Sheet', {
      pageSetup: { fitToPage: true, fitToHeight: 5, fitToWidth: 7,
        showGridLines: false, 
        printTitlesRow: '1:1',
        margins: {
          left: 0.3, right: 0.3,
          top: 0.75, bottom: 0.75,
          header: 0.1, footer: 0.1
        }
      },
      headerFooter: {
        oddHeader: `GRI - ${this.id}`,
      },

    })
    this.sheet.pageSetup.margins = {
      left: 0.7, right: 0.7,
        top: 0.75, bottom: 0.75,
          header: 0.3, footer: 0.3
    }
    this.sheet.headerFooter.oddHeader = this.id;

    this.sheet.properties.defaultRowHeight = 15
    this.sheet.columns = COLUMNS
    this.addInfo()
    const filename = "now.xlsx"
    await this.workbook.xlsx.writeFile(filename)
    return filename
  }
  addInfo(){
    this.addHeader()
    this.addColumnNames()
    this.rooms.forEach(this.addRoom.bind(this))
    this.addEmptyRow()
    this.addEmptyRow()
    this.addTotals()
    this.addEmptyRow()
    this.addEmptyRow()
    this.addFooter()
  }
  addColumnNames(){
    this.sheet.addRow(COLUMNS.map(c => c.header));
    this.sheet.lastRow.font = {
      color: { argb: 'FF444444'},
      size: 9,
    }
  }
  addHeader(){
    const { title,
      internal_id = '',
      address = '',
      clientName = '',
      details,
      provider_name,
      based_on_your_request,
    } = this.project
    const { created_at } = this.version
  
    const date = new Date(created_at).toLocaleDateString('fr')
    addImage.call(this)
    addDetails = addDetails.bind(this)
    addDetails('TVA CHE 110.257.937')
    this.addEmptyRow()
    addDetails(`OFFRE N: ${internal_id}`,'', `Genève, le ${date}` )
    this.addEmptyRow()
    addDetails(`CONCERNE: ${title}`)
    this.addEmptyRow()
    addDetails(`Adresse: ${address}`, '', `${provider_name}`)
    addDetails(`Contact: ${clientName}`)
    this.addEmptyRow()
    based_on_your_request && addDetails(`Selon votre demande de devis No: ${based_on_your_request}`)
    details && addDetails(`${details}`)

    const contact = [this.user.name, this.user.phone].filter(Boolean).join(' | ')
    addDetails(`Vos Contact: ${contact}`)
    this.addEmptyRow()
    this.addEmptyRow()

    function addDetails(...details){
      this.sheet.addRow(details)
      const lastRow = this.sheet.lastRow._number
      this.sheet.mergeCells(`A${lastRow}:B${lastRow}`)
    }
    
    function addImage(){
      var banner = this.workbook.addImage({
        buffer: fs.readFileSync('banner.png'),
        extension: 'png',
      });
      var lastRow = 5
      for(let i = 0; i < lastRow; i++){
        this.sheet.addRow()
      }
      this.sheet.addImage(banner, `A1:${LAST_COL}5`)
    }
  }
  addRoom({ name, steps, note }){
    if(!steps.length){
      return
    }
    addRoomTitle.call(this)
    if(note){
      this.sheet.addRow()
      this.sheet.addRow(['', note])
      formatCells({ ...fill('EEEEEE'), ...border(), font: { italic: true, bold: true }}, this.sheet.lastRow, ['B', 'B'])
    }
    const firstRoomRow = this.sheet.lastRow._number + 1
    let lastCategoryPrinted
    steps.forEach(addStep.bind(this))
    const lastRoomRow = this.sheet.lastRow._number
    addTotal.call(this)

    this.addEmptyRow(2)
    function addTotal(){
      const totalCol = colToLetter(this.sheet.getColumn('total')._number)
      const formula = `CEILING(SUM(${totalCol}${firstRoomRow}:${totalCol}${lastRoomRow}), 0.05)`
      const cell = this.addFormula(formula, 'Sous-total', { border: { top: 'thick', color: { argb: '#FF000' }}})
      cell.border = { top: 'thick', color: { argb: '#FF0000' } }
      this.cellsThatAreTotal.push(cell)
      formatCells(borderTop, this.sheet.lastRow, ['C', LAST_COL])
    }

    function addRoomTitle(){
      this.sheet.addRow(['', (name || 'Sans titre').toUpperCase()])
      const row = this.sheet.lastRow
      formatCells({...fill('DDDDDD'), font: { bold: true }}, row, ['B', LAST_COL])

    }

    function maybeAddStepCategory(step){
      const { category } = step
      if (lastCategoryPrinted === category){
        return
      }
      if(lastCategoryPrinted){
        this.addEmptyRow()
      }
      this.sheet.addRow(['', (category || 'Autre')])
      const row = this.sheet.lastRow
      const number = row._number
      // row.font = { bold: true }
      row._cells.forEach(c => {
        c.font = {
          bold: true,
        }
      })
      // this.sheet.mergeCells(`B${number}:${LAST_COL}${number}`)
      lastCategoryPrinted = category
    }
    function addStep(step) {
      const { 
        quantity, 
        min_price = 0, 
        description, 
        unit, 
        price 
      } = step
      maybeAddStepCategory.call(this,step)
      const stepRow = [
        this.currentStepIndex,
        description,
        parseFloat(quantity),
        unit,
        price,
        '',
        price
      ]

      this.sheet.addRow(stepRow)
      const row = this.sheet.lastRow

      const descCell = row.getCell('description')
      descCell.wrapText = true

      const priceDisplayCell = row.getCell('price_display')
      const priceCell = row.getCell('price')
      const stepTotalCell = row.getCell('total')
      const qtyCell = row.getCell('quantity')
      const isInBlocks = quantity * price <= min_price
      if(isInBlocks){
        qtyCell.value = 1
      }

      const priceFormula = `${qtyCell._address} * ${priceCell._address}`
      priceCell.font = {
        color: { argb: 'FFFFFFFF' }
      }
      // if the min_price is not reached, then we should display
      // the min_price in the price column – therefore we need to keep 
      // the step's price in the sheet
      const priceDisplayFormula = `IF(${priceFormula} <= ${min_price}, ${stepTotalCell._address}, ${priceCell._address})`
      priceDisplayCell.value = { formula: priceDisplayFormula }

      const totalFormula = `CEILING(MAX(${priceFormula}, ${min_price}), 0.05)`
      stepTotalCell.value = { formula: totalFormula }
      const blockOrUnitFormula = `IF(${priceFormula} <= ${min_price}, "bloc", "${unit}")`
      const unitCell = row.getCell('unit')
      unitCell.value = { formula: blockOrUnitFormula }

      row.eachCell({ includeEmpty: true }, c => {
        c.alignment = {
          vertical: 'top',
          wrapText: true
        }
      })
      this.currentStepIndex++
    }
  }
  addTotals(){
    const total = `CEILING(SUM(${this.cellsThatAreTotal.map(c => c._address).join(',')}), 0.05)`
    const ht = this.addFormula(total, 'TOTAL H.T')
    const tvaForm = `${ht._address}*7.7%`
    this.maybeAddDiscount()
    const tva = this.addFormula(tvaForm, 'T.V.A. 7.7%')
    this.addEmptyRow()
    const ttcForm = `${this.lastTotalCell._address}+${tva._address}`
    this.addFormula(ttcForm, 'TOTAL T.T.C.')
    formatCells(borderTop, this.sheet.lastRow, ['C', LAST_COL])
  }
  addFooter(){
    addLine = addLine.bind(this)
    if(this.isProviderPrivate()){
      // TODO only for Privé
      const terms = [
          'Ce devis a été établi sur la base des éléments dont nous disposons, ne sont pas compris tous travaux qui ne sont pas explicitement décrits.',
          "L'acceptation de ce devis implique l'entière compréhension des points énumérés.",
          "Avant tout travaux de carrelage, un test amiante est nécessaire. En cas de présence de matériaux amiantés, les mesures nécessaires devront être prises pour l'élimination de ceux-ci.",
          "Tous travaux de réfection, modification d'affectation ou de rénovations doivent être annoncés auprès du DCTI et des départements concernés. Le propriétaire s'engage à effectuer les démarches administratives lui-même ou par le biais d'un architecte. L'entreprise GRI ne serait être responsable en cas d'éventuel recours.",
      ]
      const notes = [
          "Note:",
          "Un premier acompte d'environ 30% du montant de l'adjudication sera demandé pour l'ouverture du chantier.",
          "Un deuxième acompte sera demandé à la fin du premier tiers du chantier.",
          "Un troisième acompte sera demandé à la fin du deuxième tiers du chantier.",
          "Une facture du solde du montant final sera envoyée à la fin du chantier.",
          "",
      ]
      addLine(terms.join('\n'), { font: { bold: true }, height: 100 })
      this.addEmptyRow()
      addLine(notes.join('\n'), { font: { bold: true}, height: 60 })
    }
    this.addEmptyRow()
    this.addEmptyRow()
    addLine(
      "En cas d'acceptation, nous vous remercions de nous retourner la copie de ce devis datée et signée et portant la mention manuscrite \"Bon pour accord et travaux\".",
      { font: { italic: true, size: 10 }, height: 30 }
      )
    this.addEmptyRow()
    addLine("BON POUR ACCORD", { font: { bold: true } })
    this.addEmptyRow()
    this.sheet.addRow(['', 'Signature', 'Genève, le'])

    this.addEmptyRow()
    this.addEmptyRow()
    this.addEmptyRow()
    
    addLine(
      "Selon la loi fédérale contre la concurrence déloyale, l'utilisation ou reproduction du devis sont strictement interdites sans l'autorisation écrite de l'entreprise.",
      { font: { italic: true, size: 10 }, height: 30 }
      )
    this.addEmptyRow()
    addLine("Payable à:  UBS -  IBAN CH05 0024 0240 3998 0401F")

    function addLine(text, format){
      this.sheet.addRow(['', text])
      this.merge('B', LAST_COL)
      formatCells({ ...format, ...wrap }, this.sheet.lastRow, ['B', 'B'])
    }
  }
  isProviderPrivate(){
    const { provider_name } = this.project
    return provider_name === 'Privé'
  }
  initWorkBook(){
    var workbook = new Excel.Workbook();
    workbook.creator = 'Me';
    workbook.lastModifiedBy = 'Her';
    workbook.created = new Date();
    workbook.modified = new Date();
    workbook.lastPrinted = new Date();
    this.workbook = workbook
  }
  merge(from, to){
    const lastRow = this.sheet.lastRow
    this.sheet.mergeCells(`${from}${lastRow._number}:${to}${lastRow._number}`)
    return lastRow
  }
  addEmptyRow(count = 1){
    for(let i = count; i > 0; i--){
      this.sheet.addRow()
    }
  }
  addFormula(formula, text="Total", styles = {}){
    this.sheet.addRow(['', '', '', '', text, ''])
    const row = this.sheet.lastRow
    const cell = row.getCell('total')
    cell.value = { formula };
    this.lastTotalCell = cell
    cell.font = { bold: true }
    cell.numFmt = "#,###0.##"
    row.alignment = { vertical: 'bottom', horizontal: 'right' };
    for(let key in styles){
      cell[key] = styles[key]
    }
    return cell
  }
}

//new Excelor(data).createDocument()
module.exports = Excelor
