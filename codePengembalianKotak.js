function showPengembalianKotak() {
  var formPengembalianKotak = HtmlService
    .createHtmlOutputFromFile('formPengembalianKotak')
    .setWidth(1000)
    .setHeight(750)

  SpreadsheetApp.getUi()
    .showModalDialog(formPengembalianKotak, 'Form Pengembalian Kotak')
}

function getAvailReturnKotak() {
  var user = Session.getActiveUser().getEmail()
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheetRiwayatPinjamKotak = spreadsheet.getSheetByName('Riwayat Pinjam Kotak')
  var lrRiwayatPinjamKotak = sheetRiwayatPinjamKotak.getLastRow()
  var dataRow = sheetRiwayatPinjamKotak.getRange(2,1,lrRiwayatPinjamKotak-1,6).getValues()
  var filteredData = dataRow.filter((data)=>{return !data[4] && data[5] === user})
  return filteredData
}

function saveDataPengembalianKotak(idKotakKembali, gudangTujuan) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheetPinjamKotak = spreadsheet.getSheetByName('Riwayat Pinjam Kotak')
  var lrPinjamKotak = sheetPinjamKotak.getLastRow()

  var sheetKotak = spreadsheet.getSheetByName('Kotak')
  var lrKotak = sheetKotak.getLastRow()

  var user = Session.getActiveUser().getEmail()
  var date = new Date()

  idKotakKembali.forEach(id => {
    var dataRow = sheetKotak.getRange(2,1,lrKotak-1,11).getValues().findIndex((data)=>{return data[0]===id}) + 2
    var dataPinjamRow = sheetPinjamKotak.getRange(2,1,lrPinjamKotak-1,1).getValues().findIndex((data)=>{return data[0]===id})+2
    sheetKotak.getRange(dataRow,8).setValue(gudangTujuan)
    sheetKotak.getRange(dataRow,9).setValue('Terpakai')
    sheetPinjamKotak.getRange(dataPinjamRow,5).setValue('TRUE')
    sheetPinjamKotak.getRange(dataPinjamRow,8).setValue(Utilities.formatDate(date,"Asia/Bangkok","dd MMMM yyyy HH:mm:ss"))
  });

  return 'Berhasil mencatat pengembalian kotak'
}

// function debug() {
//   console.log(getAvailReturnKotak())
// }