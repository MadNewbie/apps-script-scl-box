function showPengembalianKotakAntarGudangSementara() {
  var formPengembalianKotakAntarGudangSementara = HtmlService
    .createHtmlOutputFromFile('formPengembalianKotakAntarGudangSementara')
    .setWidth(1000)
    .setHeight(750)

    SpreadsheetApp.getUi()
      .showModalDialog(formPengembalianKotakAntarGudangSementara, 'Form Pengembalian Kotak Antar Gudang Sementara')
}

function getAvailDataKotakForPengembalianSementara(gudangAsal) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheetRiwayatPinjamSementara = spreadsheet.getSheetByName('Riwayat Pinjam Kotak Antar Gudang Sementara')
  var lrRiwayatPinjamSementara = sheetRiwayatPinjamSementara.getLastRow()
  var user = Session.getActiveUser().getEmail()

  var dataKotakDipinjam = sheetRiwayatPinjamSementara.getRange(2,1,lrRiwayatPinjamSementara-1,7).getValues()
  var res = dataKotakDipinjam.filter(data => {return data[4]===gudangAsal && !data[5] && data[6]===user})
  return res
}

function saveDataPengembalianKotakAntarGudangSementara(idKotak, gudangTujuan) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheetPinjamKotak = spreadsheet.getSheetByName('Riwayat Pinjam Kotak Antar Gudang Sementara')
  var lrPinjamKotak = sheetPinjamKotak.getLastRow()

  var sheetKotak = spreadsheet.getSheetByName('Kotak')
  var lrKotak = sheetKotak.getLastRow()

  var user = Session.getActiveUser().getEmail()
  var date = new Date()

  idKotak.forEach(id => {
    var dataRow = sheetKotak.getRange(2,1,lrKotak-1,11).getValues().findIndex((data)=>{return data[0]===id}) + 2
    var dataPinjamRow = sheetPinjamKotak.getRange(2,1,lrPinjamKotak-1,1).getValues().findIndex((data)=>{return data[0]===id})+2
    sheetKotak.getRange(dataRow,8).setValue(gudangTujuan)
    sheetKotak.getRange(dataRow,12).setValue('')
    sheetPinjamKotak.getRange(dataPinjamRow,6).setValue('TRUE')
    sheetPinjamKotak.getRange(dataPinjamRow,9).setValue(Utilities.formatDate(date,"Asia/Bangkok","dd MMMM yyyy HH:mm:ss"))
    sheetPinjamKotak.getRange(dataPinjamRow,10).setValue(gudangTujuan)
  });

  return 'Berhasil mencatat pengembalian kotak'
}