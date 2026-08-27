function showCetakLabelRange() {
  var formCetakLabelRange = HtmlService
    .createHtmlOutputFromFile('formCetakLabelRange')
    .setWidth(500)
    .setHeight(400)
    
  SpreadsheetApp.getUi()
    .showModalDialog(formCetakLabelRange, 'Form Cetak Label')
}

function clickBtnGenerateLabelRange(awal, akhir) {
  const params = {
    awal: awal,
    akhir: akhir,
  }

  PropertiesService.getUserProperties().setProperty('paramTriggerLabelRange', JSON.stringify(params))

  ScriptApp.newTrigger('printLabelBoxRange')
    .timeBased()
    .after(5*1000)
    .create()

  return 'File masih diproses dan akan dikirim melalui email. Silahkan cek email secara berkala. Sesungguhnya Tuhan membersamai orang-orang yang sabar :)'
}

function printLabelBoxRange(e) {
  console.log('fungsi print Label Range ditrigger')

  const savedProperties = PropertiesService.getUserProperties().getProperty('paramTriggerLabelRange')

  var params

  if (savedProperties) {
    params = JSON.parse(savedProperties)
  } else {
    Logger.log('trigger tidak memiliki parameter')
    return
  }

  var awal = params.awal
  var akhir = params.akhir

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheetTemplate = spreadsheet.getSheetByName('Template Label')
  var sheetKotak = spreadsheet.getSheetByName('Kotak')
  var lrKotak = sheetKotak.getLastRow()
  sheetTemplate.clear()

  //generate header
  sheetTemplate.getRange(1,1).setValue('Filling No')
  sheetTemplate.getRange(1,2).setValue('Subject')
  sheetTemplate.getRange(1,3).setValue('Area')
  sheetTemplate.getRange(1,4).setValue('KPS')
  sheetTemplate.getRange(1,5).setValue('Catatan')

  cr = 2

  var [typeBox, typeDoc, year, noAwal] = awal.split('/')
  var noAkhir = akhir.split('/')[3]
  // console.log(typeBox, typeDoc, year, noAwal, noAkhir)

  sheetKotak.getRange(2,1,lrKotak-1,11).getValues().forEach(kotak => {
    var [kotakTypeBox, kotakTypeDoc, kotakYear, kotakNo] = kotak[4].split('/')
    if(kotakTypeBox === typeBox && kotakTypeDoc === typeDoc && kotakYear === year && parseInt(kotakNo)<=parseInt(noAkhir) && parseInt(kotakNo)>=parseInt(noAwal)){
        sheetTemplate.getRange(cr,1).setValue(kotak[4])
        sheetTemplate.getRange(cr,2).setValue(kotak[2])
        sheetTemplate.getRange(cr,3).setValue(kotak[5])
        sheetTemplate.getRange(cr,4).setValue(kotak[6])
        sheetTemplate.getRange(cr,5).setValue(kotak[9])
        cr = cr + 1
    }
  });

  var dataFile = generateAndSendLink('Template Label Kotak SCL', 'Template Label')
  PropertiesService.getUserProperties().deleteProperty('paramTriggerLabelRange')

  console.log(e)
  if (e && e.triggerUid) {
    const triggerId = e.triggerUid
    deleteTriggerByUid(triggerId)
  }
  // return 'File selesai diproses dan akan dikirim melalui email. Silahkan cek email. Sesungguhnya Tuhan membersamai orang-orang yang sabar :)'

}