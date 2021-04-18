const Excel = require('exceljs');
const fs = require('fs');

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
const LAST_COL = colToLetter(COLUMNS.length - 1)

class Excelor{
  constructor(data){
    this.project = data
    this.rooms = data.version.rooms
    this.date = data.version.created_at
    this.cellsThatAreTotal = []
    this.currentStepIndex = 1
  }
  async createDocument(){
    const { id: pID, version: { vID}} = this.project
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
    const filename = "now.xls"
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
  }
  addHeader(){
    const { title,
      version: { created_at },
      agent_name = 'GRI', agent_number = "022 347 84 84",
      based_on_your_request,
    } = this.project

    const date = new Date(created_at).toLocaleDateString('fr')
    addImage.call(this)
    addDetails = addDetails.bind(this)
    addDetails('TVA CHE 110.257.937')
    this.addEmptyRow()
    addDetails(`OFFRE N: ${this.id}`,'', `Genève, le ${date}` )
    this.addEmptyRow()
    addDetails(`CONCERNE: ${title}`)
    this.addEmptyRow()
    addDetails(`Adresse: ${agent_name}`)
    addDetails(`Contact: `)
    this.addEmptyRow()
    based_on_your_request && addDetails(`Selon votre demande de devis No: ${based_on_your_request}`)
    addDetails(`Vos Contact: ${agent_number}`)
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
  addRoom({ name, steps }){
    if(!steps.length){
      return
    }
    addRoomTitle.call(this)
    const firstRoomRow = this.sheet.lastRow._number + 1
    let lastCategoryPrinted
    steps.forEach(addStep.bind(this))
    const lastRoomRow = this.sheet.lastRow._number
    addTotal.call(this)

    function addTotal(){
      const totalCol = colToLetter(this.sheet.getColumn('total')._number)
      const formula = `SUM(${totalCol}${firstRoomRow}:${totalCol}${lastRoomRow})`
      const cell = this.addFormula(formula, 'Sous-total', { border: { top: 'thick', color: '#FF000'}})
      cell.border = { top: 'thick', color: '#FF0000'}
      this.cellsThatAreTotal.push(cell)
    }

    function addRoomTitle(){
      this.sheet.addRow(['', (name || 'Sans titre').toUpperCase()])
      const row = this.sheet.lastRow
      const number = row._number
      // row.font = { bold: true }
      row._cells.forEach(c => {
        c.fill = {
          type: 'pattern',
          pattern: 'lightGray',
          fgColor: '#FF0000'
        }
      })
      this.sheet.mergeCells(`B${number}:${LAST_COL}${number}`)
    }

    function maybeAddStepCategory(step){
      const { category } = step
      if (lastCategoryPrinted === category){
        return
      }
      this.sheet.addRow(['', (category || 'Autre')])
      const row = this.sheet.lastRow
      const number = row._number
      // row.font = { bold: true }
      row._cells.forEach(c => {
        c.fill = {
          type: 'pattern',
          pattern: 'lightGray',
          fgColor: '#FF0000'
        }
      })
      this.sheet.mergeCells(`B${number}:${LAST_COL}${number}`)
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
      const priceDisplayFormula = `IF(${priceFormula} <= ${min_price}, ${stepTotalCell._address}, ${priceCell._address})`
      priceDisplayCell.value = { formula: priceDisplayFormula }

      const totalFormula = `MAX(${priceFormula}, ${min_price})`
      stepTotalCell.value = { formula: totalFormula }
      const blockOrUnitFormula = `IF(${priceFormula} <= ${min_price}, "bloc", "${unit}")`
      const unitCell = row.getCell('unit')
      unitCell.value = { formula: blockOrUnitFormula }

      row.eachCell({ includeEmpty: true }, c => {
        c.alignment = {
          vertical: 'middle',
          wrapText: true
        }
      })
      this.currentStepIndex++
    }
  }
  addTotals(){
    const total = `SUM(${this.cellsThatAreTotal.map(c => c._address).join(',')})`
    const ht = this.addFormula(total, 'TOTAL H.T')
    const tvaForm = `${ht._address}*7.7%`
    const tva = this.addFormula(tvaForm, 'T.V.A. 7.7%')
    this.addEmptyRow()
    const ttcForm = `${ht._address}+${tva._address}`
    this.addFormula(ttcForm, 'TOTAL T.T.C.')
  }
  addFooter(){
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
    addLine = addLine.bind(this)
    addLine(terms.join('\n'), { font: { bold: true }, height: 100 })
    this.addEmptyRow()
    addLine(notes.join('\n'), { font: { bold: true}, height: 60 })
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
      const lastRow = this.sheet.lastRow
      this.merge('B', LAST_COL)
      for(let key in format){
        lastRow[key] = format[key]
      }
      lastRow.alignment = { wrapText: true }
      return lastRow
    }
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
  addEmptyRow(){
    this.sheet.addRow()
  }
  addFormula(formula, text="Total", styles = {}){
    this.sheet.addRow(['', '', '', '', text, ''])
    const row = this.sheet.lastRow
    const cell = row.getCell('total')
    cell.value = { formula };
    cell.font = { bold: true }
    row.alignment = { vertical: 'bottom', horizontal: 'right' };
    for(let key in styles){
      cell[key] = styles[key]
    }
    return cell
  }
}

//new Excelor(data).createDocument()
module.exports = Excelor
