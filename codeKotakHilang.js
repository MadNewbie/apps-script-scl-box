function showKotakHilang() {
  var formKotakHilang = HtmlService
    .createHtmlOutputFromFile('formKotakHilang')
    .setWidth(1000)
    .setHeight(750)
    
  SpreadsheetApp.getUi()
    .showModalDialog(formKotakHilang, 'Form Kotak Hilang')
}

function getAvailDataKotakHilang(gudang) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
	var sheetKotak = spreadsheet.getSheetByName('Kotak')
	var lrKotak = sheetKotak.getLastRow()
	var dataRow = sheetKotak.getRange(2,1,lrKotak-1,11).getValues()
	var filteredData = dataRow.filter((data)=>{return (data[7]===gudang && data[8]!=='Hilang')})
	return filteredData
}

function saveDataKotakHilang(idKotakHilang, kronologiHilang) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()

  var sheetKotak = spreadsheet.getSheetByName('Kotak')
  var lrKotak = sheetKotak.getLastRow()

  var sheetKotakHilang = spreadsheet.getSheetByName('Riwayat Kotak Hilang')
  var lrKotakHilang = sheetKotakHilang.getLastRow()

  var user = Session.getActiveUser().getEmail()
  var date = new Date()

  idKotakHilang.forEach((id) => {
    var dataRow = sheetKotak.getRange(2,1,lrKotak-1,11).getValues().findIndex((data)=>{return data[0]===id})+2
    sheetKotak.getRange(dataRow,9).setValue('Hilang')
    sheetKotakHilang.getRange(lrKotakHilang+1,1).setValue(sheetKotak.getRange(dataRow,1).getValue())
    sheetKotakHilang.getRange(lrKotakHilang+1,2).setValue(sheetKotak.getRange(dataRow,5).getValue())
    sheetKotakHilang.getRange(lrKotakHilang+1,3).setValue(sheetKotak.getRange(dataRow,8).getValue())
    sheetKotakHilang.getRange(lrKotakHilang+1,4).setValue(user)
    sheetKotakHilang.getRange(lrKotakHilang+1,5).setValue(Utilities.formatDate(date,"Asia/Bangkok","dd MMMM yyyy HH:mm:ss"))
    sheetKotakHilang.getRange(lrKotakHilang+1,6).setValue(kronologiHilang)
    lrKotakHilang = lrKotakHilang + 1
  })
  return 'Berhasil mencatat data kotak hilang'
}