function showCetakLabel() {
  var formCetakLabel = HtmlService
    .createHtmlOutputFromFile('formCetakLabel')
    .setWidth(300)
    .setHeight(200)
    
  SpreadsheetApp.getUi()
    .showModalDialog(formCetakLabel, 'Form Cetak Label')
}

function getAvailDataCetakKotak(jenisDokumen) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
	var sheetKotak = spreadsheet.getSheetByName('Kotak')
	var lrKotak = sheetKotak.getLastRow()
	var dataRow = sheetKotak.getRange(2,1,lrKotak-1,11).getValues()
	var filteredData = dataRow.filter((data)=>{
    return (data[2]===jenisDokumen && data[8]==='Terpakai')
  })
	return filteredData
}

function clickBtnGenerateLabel(idCetakKotak) {
  const params = {
    idCetakKotak: idCetakKotak
  }

  PropertiesService.getUserProperties().setProperty('paramTriggerLabel', JSON.stringify(params))

  ScriptApp.newTrigger('printLabelBox')
    .timeBased()
    .after(5*1000)
    .create()

  return 'File masih diproses dan akan dikirim melalui email. Silahkan cek email secara berkala. Sesungguhnya Tuhan membersamai orang-orang yang sabar :)'
}

function printLabelBox(e) {

  console.log('fungsi print Label ditrigger')

  const savedProperties = PropertiesService.getUserProperties().getProperty('paramTriggerLabel')

  var params

  if (savedProperties) {
    params = JSON.parse(savedProperties)
  } else {
    Logger.log('trigger tidak memiliki parameter')
    return
  }

  var idCetakKotak = params.idCetakKotak

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

  idCetakKotak.forEach(id => {
    var dataRow = sheetKotak.getRange(2,1,lrKotak-1,11).getValues().findIndex((data)=>{return data[0]===id})+2
    sheetTemplate.getRange(cr,1).setValue(sheetKotak.getRange(dataRow,5).getValue())
    sheetTemplate.getRange(cr,2).setValue(sheetKotak.getRange(dataRow,3).getValue())
    sheetTemplate.getRange(cr,3).setValue(sheetKotak.getRange(dataRow,6).getValue())
    sheetTemplate.getRange(cr,4).setValue(sheetKotak.getRange(dataRow,7).getValue())
    sheetTemplate.getRange(cr,5).setValue(sheetKotak.getRange(dataRow,10).getValue())
    cr = cr + 1
  });

  var dataFile = generateAndSendLink('Template Label Kotak SCL', 'Template Label')
  PropertiesService.getUserProperties().deleteProperty('paramTriggerLabel')

  console.log(e)
  if (e && e.triggerUid) {
    const triggerId = e.triggerUid
    deleteTriggerByUid(triggerId)
  }
  // return 'File selesai diproses dan akan dikirim melalui email. Silahkan cek email. Sesungguhnya Tuhan membersamai orang-orang yang sabar :)'

}