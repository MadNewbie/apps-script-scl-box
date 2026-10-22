function showTambahKotakBaruForm() {
  var formTambahKotakBaru = HtmlService
    .createHtmlOutputFromFile('formTambahKotakBaru')
    .setWidth(300)
    .setHeight(300)

  SpreadsheetApp.getUi()
    .showModalDialog(formTambahKotakBaru, 'Form Tambah Kotak Baru')
}

function saveDataTambahKotakBaru(formData) {
  const action = 'kbt'
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

function kotakBaruTambah(param) {
  var spreadsheet = SpreadsheetApp.openById(param.fileId)

  if (!spreadsheet) {
    return jsonResponse({
      'code': 404,
      'message': "File tidak ditemukan",
    })
  }

  var sheet = spreadsheet.getSheetByName('Riwayat Kotak Baru')
  
  if (!sheet) {
    return jsonResponse({
      'code': 404,
      'message': "Sheet tidak ditemukan",
    })
  }

  var lastRow = sheet.getLastRow()
  var user = param.user

  sheet.getRange(lastRow+1, 1).setValue(param.jenisKotak)
  sheet.getRange(lastRow+1, 2).setValue(param.jumlahKotak)
  sheet.getRange(lastRow+1, 3).setValue(new Date())
  sheet.getRange(lastRow+1, 4).setValue(param.gudang)
  sheet.getRange(lastRow+1, 5).setValue(user)

  return jsonResponse({
    'code':200,
    'message':'Berhasil memasukkan riwayat kotak baru'
  })
}
