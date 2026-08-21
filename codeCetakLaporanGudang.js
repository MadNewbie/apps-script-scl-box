function showFormCetakLaporanGudang() {
  var formCetakLaporanGudang = HtmlService
    .createHtmlOutputFromFile('formCetakLaporanGudang')
    .setWidth(300)
    .setHeight(200)

  SpreadsheetApp.getUi()
    .showModalDialog(formCetakLaporanGudang, 'Form Cetak Laporan Gudang')
}

function printReportGudang(gudang) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheetKotak = spreadsheet.getSheetByName('Kotak')
  var sheetTemplate = spreadsheet.getSheetByName('Report Template')
  var lrSheetKotak = sheetKotak.getLastRow()
  var date = new Date()
  sheetTemplate.clear()

    //generate header
  sheetTemplate.getRange(1,1).setValue('Laporan KPS Retail').setFontWeight('bold')
  sheetTemplate.getRange(2,1).setValue('Tanggal Cetak').setFontWeight('bold')
  sheetTemplate.getRange(2,2).setValue(Utilities.formatDate(date,"Asia/Bangkok","dd MMMM yyyy HH:mm:ss"))
  sheetTemplate.getRange(3,1).setValue('Label Kotak').setFontWeight('bold')
  sheetTemplate.getRange(3,2).setValue('Area').setFontWeight('bold')
  sheetTemplate.getRange(3,3).setValue('Jenis Dokumen').setFontWeight('bold')
  sheetTemplate.getRange(3,4).setValue('KPS').setFontWeight('bold')
  sheetTemplate.getRange(3,5).setValue('Catatan').setFontWeight('bold')
  sheetTemplate.getRange(3,6).setValue('Status Khusus').setFontWeight('bold')

  cr = 4

  sheetKotak.getRange(2,1,lrSheetKotak-1,12).getValues().filter((kotak) => {
    return kotak[7]===gudang && kotak[8]==='Terpakai'
  }).forEach(kotak => {
    console.log(kotak)
    sheetTemplate.getRange(cr,1).setValue(kotak[4])
    sheetTemplate.getRange(cr,2).setValue(kotak[5])
    sheetTemplate.getRange(cr,3).setValue(kotak[2])
    sheetTemplate.getRange(cr,4).setValue(kotak[6])
    sheetTemplate.getRange(cr,5).setValue(kotak[10])
    sheetTemplate.getRange(cr,6).setValue(kotak[11])
    cr++
  })

  var dataFile = generateAndSendLink(`Laporan Data Gudang Per Tanggal - ${Utilities.formatDate(date,"Asia/Bangkok","dd MMMM yyyy")}`)
  return 'File selesai diproses dan akan dikirim melalui email. Silahkan cek email. Sesungguhnya Tuhan membersamai orang-orang yang sabar :)'
}