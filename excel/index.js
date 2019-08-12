
function addHeaders(worksheet){
  worksheet.columns = [
    { header: 'Id', key: 'id', width: 10 },
    { header: 'Name', key: 'name', width: 32 },
    { header: 'D.O.B.', key: 'DOB', width: 10, outlineLevel: 1 }
  ];
}

const excelor = async () => {
  const Excel = require('exceljs/modern.nodejs');
  var workbook = new Excel.Workbook();

  workbook.creator = 'Me';
  workbook.lastModifiedBy = 'Her';
  workbook.created = new Date(1985, 8, 30);
  workbook.modified = new Date();
  workbook.lastPrinted = new Date(2016, 9, 27);


  var sheet = workbook.addWorksheet('My Sheet');
  addHeaders(sheet)
  sheet.addRow([4, 'Sam', new Date()]);

  const filename = "now.xls"
  await workbook.xlsx.writeFile(filename)

  // write to a stream
  // await workbook.xlsx.write(stream)

  // // write to a new buffer
  // workbook.xlsx.writeBuffer()
  //   .then(function (buffer) {
  //     // done
  //   });
  return filename
}

module.exports = excelor