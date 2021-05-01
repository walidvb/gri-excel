const discounts = [
  {
    text: 'Rabais -3%',
    discount: '-3%',
  },
  {
    text: 'Escompte -2%',
    discount: '-2%',
  },
  {
    text: 'Prorata -1.5%',
    discount: '-1.5%',
  },
]

function maybeAddDiscount(){
  if(this.isProviderPrivate()){
    return
  }
  
  discounts.forEach(({ text, discount }) => {
    this.sheet.addRow()
    const row = this.sheet.lastRow
    row.getCell('price_display').value = text
    const discountCell = row.getCell('total')
    discountCell.value = discount
    const formula = `CEILING(${this.lastTotalCell._address} - ${this.lastTotalCell._address}*${discountCell._address}, 0.05)`
    this.addFormula(formula, 'NOUVEAU SOUS TOTAL HT')
    this.sheet.addRow()
  })
  
}

module.exports = maybeAddDiscount