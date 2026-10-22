function showFormKotakBaruHilang() {
  var formKotakBaruHilang = HtmlService
    .createHtmlOutputFromFile('formKotakBaruHilang')
    .setWidth(500)
    .setHeight(400)

  SpreadsheetApp.getUi()
    .showModalDialog(formKotakBaruHilang, 'Form Kotak Baru Hilang')
}

function saveDataKotakBaruHilang(formData) {
  const action = 'kbh'
  const data = {...formData}
  const permission = checkPermission(action)

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

function kotakBaruHilang(param) {
  let spreadsheet = SpreadsheetApp.openById(param.fileId)

  if (!spreadsheet) {
    return jsonResponse({
      'code': 404,
      'message': 'File tidak ditemukan'
    })
  }

  let sheet = spreadsheet.getSheetByName('Riwayat Kotak Baru Hilang')

  if (!sheet) {
    return jsonResponse({
      'code': 404,
      'message': "Sheet tidak ditemukan",
    })
  }

  let lastRow = sheet.getLastRow()
  let user = param.user
  let date = new Date()

  sheet.getRange(lastRow + 1, 1).setValue(Utilities.formatDate(date, 'Asia/Bangkok', 'dd MMMM yyyy HH:mm:ss'))
  sheet.getRange(lastRow + 1, 2).setValue(param.jenisKotak)
  sheet.getRange(lastRow + 1, 3).setValue(param.jumlahKotak)
  sheet.getRange(lastRow + 1, 4).setValue(param.gudang)
  sheet.getRange(lastRow + 1, 5).setValue(param.kronologi)
  sheet.getRange(lastRow + 1, 6).setValue(user)

  return jsonResponse({
    'code':200,
    'message':'Berhasil memasukkan riwayat kotak baru'
  })
}