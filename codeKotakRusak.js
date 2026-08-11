function showKotakRusak() {
  var formKotakRusak = HtmlService
    .createHtmlOutputFromFile('formKotakRusak')
    .setWidth(1000)
    .setHeight(750)

  SpreadsheetApp.getUi()
    .showModalDialog(formKotakRusak, 'Form Kotak Rusak')
}

function getAvailDataKotakRusak(gudang) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
	var sheetKotak = spreadsheet.getSheetByName('Kotak')
	var lrKotak = sheetKotak.getLastRow()
	var dataRow = sheetKotak.getRange(2,1,lrKotak-1,11).getValues()
	var filteredData = dataRow.filter((data)=>{return (data[7]===gudang && data[8]!=='Hilang' && data[8]!=='Rusak')})
	return filteredData
}

function saveDataKotakRusak(idKotakRusak, kronologiRusak) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()

  var sheetKotak = spreadsheet.getSheetByName('Kotak')
  var lrKotak = sheetKotak.getLastRow()

  var sheetKotakRusak = spreadsheet.getSheetByName('Riwayat Kotak Rusak')
  var lrKotakRusak = sheetKotakRusak.getLastRow()

  var sheetJmlKotak = spreadsheet.getSheetByName('Jumlah Kotak')
  var lrJmlKotak = sheetJmlKotak.getLastRow()

  var user = Session.getActiveUser().getEmail()
  var date = new Date()

  idKotakRusak.forEach((id) => {
    var dataRow = sheetKotak.getRange(2,1,lrKotak-1,11).getValues().findIndex((data)=>{return data[0]===id})+2
    sheetKotak.getRange(dataRow,9).setValue('Rusak')

    //menambah kotak baru
    rowJenisKotak = sheetJmlKotak.getRange(1,1,lrJmlKotak,1).getValues().findIndex(row=>row[0]===sheetKotak.getRange(dataRow,2).getValue())+1
    nextNo = sheetJmlKotak.getRange(rowJenisKotak,2).getValue() - sheetJmlKotak.getRange(rowJenisKotak,6).getValue() + 1

    sheetKotak.getRange(lrKotak+1,1).setValue(`${sheetKotak.getRange(dataRow,2).getValue()}.${nextNo}`)
    sheetKotak.getRange(lrKotak+1,2).setValue(sheetKotak.getRange(dataRow,2).getValue())
    sheetKotak.getRange(lrKotak+1,3).setValue(sheetKotak.getRange(dataRow,3).getValue())
    sheetKotak.getRange(lrKotak+1,4).setValue(sheetKotak.getRange(dataRow,4).getValue())
    sheetKotak.getRange(lrKotak+1,5).setValue(sheetKotak.getRange(dataRow,5).getValue())
    sheetKotak.getRange(lrKotak+1,6).setValue(sheetKotak.getRange(dataRow,6).getValue())
    sheetKotak.getRange(lrKotak+1,7).setValue(sheetKotak.getRange(dataRow,7).getValue())
    sheetKotak.getRange(lrKotak+1,8).setValue(sheetKotak.getRange(dataRow,8).getValue())
    sheetKotak.getRange(lrKotak+1,9).setValue('Terpakai')
    sheetKotak.getRange(lrKotak+1,10).setValue(sheetKotak.getRange(dataRow,10).getValue())
    sheetKotak.getRange(lrKotak+1,11).setValue(sheetKotak.getRange(dataRow,11).getValue())
    sheetKotak.getRange(lrKotak+1,12).setValue(sheetKotak.getRange(dataRow,12).getValue())
    sheetKotak.getRange(lrKotak+1,13).setValue(sheetKotak.getRange(dataRow,13).getValue())
    sheetKotak.getRange(lrKotak+1,14).setValue(Utilities.formatDate(date,"Asia/Bangkok","dd MMMM yyyy HH:mm:ss"))
    sheetKotak.getRange(lrKotak+1,15).setValue(user)

    //mencatat pada sheet riwayat kotak rusak
    sheetKotakRusak.getRange(lrKotakRusak+1,1).setValue(sheetKotak.getRange(dataRow,1).getValue())
    sheetKotakRusak.getRange(lrKotakRusak+1,2).setValue(sheetKotak.getRange(lrKotak+1,1).getValue())
    sheetKotakRusak.getRange(lrKotakRusak+1,3).setValue(sheetKotak.getRange(dataRow,5).getValue())
    sheetKotakRusak.getRange(lrKotakRusak+1,4).setValue(sheetKotak.getRange(dataRow,8).getValue())
    sheetKotakRusak.getRange(lrKotakRusak+1,5).setValue(user)
    sheetKotakRusak.getRange(lrKotakRusak+1,6).setValue(Utilities.formatDate(date,"Asia/Bangkok","dd MMMM yyyy HH:mm:ss"))
    sheetKotakRusak.getRange(lrKotakRusak+1,7).setValue(kronologiRusak)
    lrKotak = lrKotak + 1
    lrKotakRusak = lrKotakRusak + 1
  })
  return 'Berhasil mencatat data kotak rusak'
}