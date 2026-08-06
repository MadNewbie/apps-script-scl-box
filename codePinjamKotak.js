function showPinjamKotakForm() {
	var formPinjamKotak = HtmlService
		.createHtmlOutputFromFile('formPinjamKotak')
		.setHeight(1750)
		.setWidth(2000)

	SpreadsheetApp.getUi()
		.showModalDialog(formPinjamKotak, 'Form Pinjam Kotak')
}

function getAvailDataKotak(jenisDokumen, gudang) {
	var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
	var sheetKotak = spreadsheet.getSheetByName('Kotak')
	var lrKotak = sheetKotak.getLastRow()
	var dataRow = sheetKotak.getRange(2,1,lrKotak-1,11).getValues()
	var filteredData = dataRow.filter((data)=>{return (data[2]===jenisDokumen && data[7]===gudang && data[8]==='Terpakai')})
	return filteredData
}

function saveDataPinjamKotak(idKotakDipinjam, tujuanPeminjaman) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheetPinjamKotak = spreadsheet.getSheetByName('Riwayat Pinjam Kotak')
  var lrPinjamKotak = sheetPinjamKotak.getLastRow()

  var sheetKotak = spreadsheet.getSheetByName('Kotak')
  var lrKotak = sheetKotak.getLastRow()

  var user = Session.getActiveUser().getEmail()
  var date = new Date()

  idKotakDipinjam.forEach(id => {
    var dataRow = sheetKotak.getRange(2,1,lrKotak-1,11).getValues().findIndex((data)=>{return data[0]===id}) + 2
    sheetKotak.getRange(dataRow,9).setValue('Dipinjam')
    sheetPinjamKotak.getRange(lrPinjamKotak+1,1).setValue(sheetKotak.getRange(dataRow,5).getValue())
    sheetPinjamKotak.getRange(lrPinjamKotak+1,2).setValue(tujuanPeminjaman)
    sheetPinjamKotak.getRange(lrPinjamKotak+1,3).setValue(sheetKotak.getRange(dataRow,8).getValue())
    sheetPinjamKotak.getRange(lrPinjamKotak+1,4).setValue('FALSE')
    sheetPinjamKotak.getRange(lrPinjamKotak+1,5).setValue(user)
    sheetPinjamKotak.getRange(lrPinjamKotak+1,6).setValue(Utilities.formatDate(date,"Asia/Bangkok","dd MMMM yyyy HH:mm:ss"))
  });

  return 'Berhasil mencatat peminjaman kotak'
}
// function debug() {
//   console.log('mulai debug')
// 	var jenisDokumen = 'Retail'
// 	var gudang = 'ESA Sampoerna'
// 	var res = getAvailDataKotak(jenisDokumen, gudang)
// 	console.log(res)
// }