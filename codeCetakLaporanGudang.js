function showFormCetakLaporanGudang() {
  var formCetakLaporanGudang = HtmlService
    .createHtmlOutputFromFile('formCetakLaporanGudang')
    .setWidth(300)
    .setHeight(200)

  SpreadsheetApp.getUi()
    .showModalDialog(formCetakLaporanGudang, 'Form Cetak Laporan Gudang')
}

function clickBtnGenerateReportWarehouse(gudang) {
  const params = {
    gudang: gudang
  }

  PropertiesService.getUserProperties().setProperty('paramTriggerReportWarehouse', JSON.stringify(params))

  ScriptApp.newTrigger('printReportGudang')
    .timeBased()
    .after(5*1000)
    .create()

  return 'File masih diproses dan akan dikirim melalui email. Silahkan cek email secara berkala. Sesungguhnya Tuhan membersamai orang-orang yang sabar :)'
}

function printReportGudang(e) {
  console.log('fungsi print Report Warehouse ditrigger')

  const savedProperties = PropertiesService.getUserProperties().getProperty('paramTriggerReportWarehouse')

  var params

  if (savedProperties) {
    params = JSON.parse(savedProperties)
  } else {
    Logger.log('trigger tidak memiliki parameter')
    return
  }

  var gudang = params.gudang
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheetKotak = spreadsheet.getSheetByName('Kotak')
  var sheetTemplate = spreadsheet.getSheetByName('Template Warehouse')
  var lrSheetKotak = sheetKotak.getLastRow()
  var date = new Date()
  sheetTemplate.clear()

    //generate header
  sheetTemplate.getRange(1,1).setValue('Laporan Per Gudang').setFontWeight('bold')
  sheetTemplate.getRange(2,1).setValue('Gudang').setFontWeight('bold')
  sheetTemplate.getRange(2,2).setValue(gudang)
  sheetTemplate.getRange(3,1).setValue('Tanggal Cetak').setFontWeight('bold')
  sheetTemplate.getRange(3,2).setValue(Utilities.formatDate(date,"Asia/Bangkok","dd MMMM yyyy HH:mm:ss"))
  sheetTemplate.getRange(5,1).setValue('Label Kotak').setFontWeight('bold')
  sheetTemplate.getRange(5,2).setValue('Area').setFontWeight('bold')
  sheetTemplate.getRange(5,3).setValue('Jenis Dokumen').setFontWeight('bold')
  sheetTemplate.getRange(5,4).setValue('KPS').setFontWeight('bold')
  sheetTemplate.getRange(5,5).setValue('Periode').setFontWeight('bold')
  sheetTemplate.getRange(5,6).setValue('Tanggal Mulai').setFontWeight('bold')
  sheetTemplate.getRange(5,7).setValue('Tanggal Akhir').setFontWeight('bold')
  sheetTemplate.getRange(5,8).setValue('Catatan').setFontWeight('bold')
  sheetTemplate.getRange(5,9).setValue('Status Khusus').setFontWeight('bold')

  cr = 6

  var objKotak = getAllDataAsObject()
  // console.log(objKotak)
  var filtered = objKotak.filter((kotak) => {
    // console.log(kotak)
    return kotak.lokasi === gudang && kotak.status==='Terpakai'
  })
  filtered.forEach(kotak => {
    for (const periode of kotak.periode) {
      if (kotak.jenisDokumen !== 'Consent Letter') {
        const periodeDiKotak = Object.keys(periode)[0]
        const kpsDiKotak = Object.values(periode)[0]
        // console.log(kpsDiKotak)
        kotak.area.forEach(area => {
          kpsDiKotak.forEach(kps => {
            var kpsDate = getKpsDate(parseInt(kps),parseInt(periodeDiKotak))
            sheetTemplate.getRange(cr,1).setValue(kotak.labelKotak)
            sheetTemplate.getRange(cr,2).setValue(area)
            sheetTemplate.getRange(cr,3).setValue(kotak.jenisDokumen)
            sheetTemplate.getRange(cr,4).setValue(kps)
            sheetTemplate.getRange(cr,5).setValue(periodeDiKotak)
            sheetTemplate.getRange(cr,6).setValue(kpsDate.startDate).setNumberFormat("dd-MMM-yyyy")
            sheetTemplate.getRange(cr,7).setValue(kpsDate.endDate).setNumberFormat("dd-MMM-yyyy")
            sheetTemplate.getRange(cr,8).setValue(kotak.catatan)
            sheetTemplate.getRange(cr,9).setValue(kotak.statusKhusus)
            cr++
          })
        })
      } else {
        sheetTemplate.getRange(cr,1).setValue(kotak.labelKotak)
        sheetTemplate.getRange(cr,2).setValue(kotak.area.join(', '))
        sheetTemplate.getRange(cr,3).setValue(kotak.jenisDokumen)
        sheetTemplate.getRange(cr,4).setValue('-')
        sheetTemplate.getRange(cr,5).setValue(periode)
        // sheetTemplate.getRange(cr,6).setValue(Utilities(kpsDate.startDate,"Asia/Bangkok", "dd MMMM yyyy"))
        // sheetTemplate.getRange(cr,7).setValue(Utilities(kpsDate.endDate,"Asia/Bangkok", "dd MMMM yyyy"))
        sheetTemplate.getRange(cr,8).setValue(kotak.catatan)
        sheetTemplate.getRange(cr,9).setValue(kotak.statusKhusus)
        cr++
      }
    }
    // sheetTemplate.getRange(cr,4).setValue(kotak.periode)
    // sheetTemplate.getRange(cr,7).setValue(kotak[10])
    // sheetTemplate.getRange(cr,8).setValue(kotak[11])
  })

  var dataFile = generateAndSendLink(`Laporan Data Gudang Per Tanggal - ${Utilities.formatDate(date,"Asia/Bangkok","dd MMMM yyyy")}`, 'Template Warehouse')
  PropertiesService.getUserProperties().deleteProperty('paramTriggerReportWarehouse')

  console.log(e)
  if (e && e.triggerUid) {
    const triggerId = e.triggerUid
    deleteTriggerByUid(triggerId)
  }
  // return 'File selesai diproses dan akan dikirim melalui email. Silahkan cek email. Sesungguhnya Tuhan membersamai orang-orang yang sabar :)'

}