function showPinjamKotakAntarGudang() {
  var formPinjamKotakAntarGudang = HtmlService
    .createHtmlOutputFromFile('formPinjamKotakAntarGudang')
    .setWidth(1000)
    .setHeight(750)

    SpreadsheetApp.getUi()
      .showModalDialog(formPinjamKotakAntarGudang, 'Form Pinjam Kotak Antar Gudang')
}

function saveDataPinjamKotakAntarGudang(idKotakDipinjam, tujuanPeminjaman, gudangTujuan) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheetPinjamKotak = spreadsheet.getSheetByName('Riwayat Pinjam Kotak Antar Gudang')
  var lrPinjamKotak = sheetPinjamKotak.getLastRow()

  var sheetKotak = spreadsheet.getSheetByName('Kotak')
  var lrKotak = sheetKotak.getLastRow()

  var user = Session.getActiveUser().getEmail()
  var date = new Date()

  idKotakDipinjam.forEach(id => {
    var dataRow = sheetKotak.getRange(2,1,lrKotak-1,11).getValues().findIndex((data)=>{return data[0]===id}) + 2
    sheetPinjamKotak.getRange(lrPinjamKotak+1,1).setValue(sheetKotak.getRange(dataRow,5).getValue())
    sheetPinjamKotak.getRange(lrPinjamKotak+1,2).setValue(tujuanPeminjaman)
    sheetPinjamKotak.getRange(lrPinjamKotak+1,3).setValue(sheetKotak.getRange(dataRow,8).getValue())
    sheetPinjamKotak.getRange(lrPinjamKotak+1,4).setValue(gudangTujuan)
    sheetPinjamKotak.getRange(lrPinjamKotak+1,5).setValue('FALSE')
    sheetPinjamKotak.getRange(lrPinjamKotak+1,6).setValue(user)
    sheetPinjamKotak.getRange(lrPinjamKotak+1,7).setValue(Utilities.formatDate(date,"Asia/Bangkok","dd MMMM yyyy HH:mm:ss"))
    sheetKotak.getRange(dataRow,8).setValue(gudangTujuan)
    sheetKotak.getRange(dataRow,9).setValue('Dipinjam')
    lrPinjamKotak = lrPinjamKotak + 1
  });

  return 'Berhasil mencatat peminjaman kotak'
}