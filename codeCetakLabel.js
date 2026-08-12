function showCetakLabel() {
  var formCetakLabel = HtmlService
    .createHtmlOutputFromFile('formCetakLabel')
    .setWidth(1000)
    .setHeight(750)
    
  SpreadsheetApp.getUi()
    .showModalDialog(formCetakLabel, 'Form Cetak Label')
}

function getAvailDataCetakKotak(gudang) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
	var sheetKotak = spreadsheet.getSheetByName('Kotak')
	var lrKotak = sheetKotak.getLastRow()
	var dataRow = sheetKotak.getRange(2,1,lrKotak-1,11).getValues()
	var filteredData = dataRow.filter((data)=>{
    return (data[7]===gudang && data[8]==='Terpakai')
  })
	return filteredData
}

function printLabelBox(idCetakKotak) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheetTemplate = spreadsheet.getSheetByName('Report Template')
  var sheetKotak = spreadsheet.getSheetByName('Kotak')
  var lrKotak = sheetKotak.getLastRow()

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

  var dataFile = generateFileAndGetUrl('Template Label Kotak SCL')
  return dataFile
}