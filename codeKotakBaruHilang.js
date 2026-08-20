function showFormKotakBaruHilang() {
  var formKotakBaruHilang = HtmlService
    .createHtmlOutputFromFile('formKotakBaruHilang')
    .setWidth(500)
    .setHeight(400)

  SpreadsheetApp.getUi()
    .showModalDialog(formKotakBaruHilang, 'Form Kotak Baru Hilang')
}

function saveDataKotakBaruHilang(formData) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = spreadsheet.getSheetByName('Riwayat Kotak Baru Hilang')
  var lastRow = sheet.getLastRow()
  var user = Session.getActiveUser().getEmail()
  var date = new Date()

  sheet.getRange(lastRow + 1, 1).setValue(Utilities.formatDate(date, 'Asia/Bangkok', 'dd MMMM yyyy HH:mm:ss'))
  sheet.getRange(lastRow + 1, 2).setValue(formData.jenisKotak)
  sheet.getRange(lastRow + 1, 3).setValue(formData.jumlahKotak)
  sheet.getRange(lastRow + 1, 4).setValue(formData.gudang)
  sheet.getRange(lastRow + 1, 5).setValue(formData.kronologi)
  sheet.getRange(lastRow + 1, 6).setValue(user)
  return 'Berhasil mencatat data kotak hilang'
}