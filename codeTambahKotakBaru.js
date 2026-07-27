function showTambahKotakBaruForm() {
  var formTambahKotakBaru = HtmlService
    .createHtmlOutputFromFile('formTambahKotakBaru')
    .setWidth(300)
    .setHeight(300)

  SpreadsheetApp.getUi()
    .showModalDialog(formTambahKotakBaru, 'Form Tambah Kotak Baru')
}

function saveDataTambahKotakBaru(formData) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = spreadsheet.getSheetByName('Riwayat Kotak Baru')
  var lastRow = sheet.getLastRow()
  var user = Session.getActiveUser().getEmail()

  sheet.getRange(lastRow+1, 1).setValue(formData.jenisKotak)
  sheet.getRange(lastRow+1, 2).setValue(formData.jumlahKotak)
  sheet.getRange(lastRow+1, 3).setValue(new Date())
  sheet.getRange(lastRow+1, 4).setValue(formData.gudang)
  sheet.getRange(lastRow+1, 5).setValue(user)

  return 'Berhasil memasukkan riwayat kotak baru'
}
