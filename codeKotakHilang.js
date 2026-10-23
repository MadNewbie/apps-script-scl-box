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
	var filteredData = dataRow.filter((data)=>{return (data[7]===gudang && data[8]!=='Hilang' && data[8]!=='Rusak')})
	return filteredData
}

function saveDataKotakHilang(idKotakHilang, kronologiHilang) {
  const action = 'khg'
  let data = {}
  const permission = checkPermission(action)

  if (!permission.access) {
    return permission.message
  }

  //proses kotak hilang
  data.idKotakHilang = idKotakHilang.join(",")
  data.kronologiHilang = kronologiHilang
  data.fileId = SpreadsheetApp.getActiveSpreadsheet().getId()
  data.user = permission.user
  data.action = action

  const responseFromBE = sendDataToBE(data)

  if (responseFromBE.success) {
    return responseFromBE.message
  } else {
    throw new Error(responseFromBE.message)
  }
}

function kotakHilang(param) {
  var spreadsheet = SpreadsheetApp.openById(param.fileId)

  if (!spreadsheet) {
    return jsonResponse({
      'code': 404,
      'message': "File tidak ditemukan",
    })
  }

  var sheetKotak = spreadsheet.getSheetByName('Kotak')

  if (!sheetKotak) {
    return jsonResponse({
      'code': 404,
      'message': "Sheet tidak ditemukan",
    })
  }

  var lrKotak = sheetKotak.getLastRow()

  var sheetKotakHilang = spreadsheet.getSheetByName('Riwayat Kotak Hilang')

  if (!sheetKotakHilang) {
    return jsonResponse({
      'code': 404,
      'message': "Sheet tidak ditemukan",
    })
  }

  var lrKotakHilang = sheetKotakHilang.getLastRow()

  var user = param.user
  var date = new Date()

  const idKotakHilang = param.idKotakHilang.split(",")

  idKotakHilang.forEach((id) => {
    var dataRow = sheetKotak.getRange(2,1,lrKotak-1,11).getValues().findIndex((data)=>{return data[0]===id})+2
    sheetKotak.getRange(dataRow,9).setValue('Hilang')
    sheetKotakHilang.getRange(lrKotakHilang+1,1).setValue(sheetKotak.getRange(dataRow,1).getValue())
    sheetKotakHilang.getRange(lrKotakHilang+1,2).setValue(sheetKotak.getRange(dataRow,5).getValue())
    sheetKotakHilang.getRange(lrKotakHilang+1,3).setValue(sheetKotak.getRange(dataRow,8).getValue())
    sheetKotakHilang.getRange(lrKotakHilang+1,4).setValue(user)
    sheetKotakHilang.getRange(lrKotakHilang+1,5).setValue(Utilities.formatDate(date,"Asia/Bangkok","dd MMMM yyyy HH:mm:ss"))
    sheetKotakHilang.getRange(lrKotakHilang+1,6).setValue(param.kronologiHilang)
    lrKotakHilang = lrKotakHilang + 1
  })

  return jsonResponse({
    'code': 200,
    'message': "Berhasil mencatat data kotak hilang",
  })
}