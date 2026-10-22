function showPinjamKotakAntarGudang() {
  var formPinjamKotakAntarGudang = HtmlService
    .createHtmlOutputFromFile('formPinjamKotakAntarGudang')
    .setWidth(1000)
    .setHeight(750)

    SpreadsheetApp.getUi()
      .showModalDialog(formPinjamKotakAntarGudang, 'Form Transfer Kotak Antar Gudang')
}

function saveDataPinjamKotakAntarGudang(idKotakDipinjam, tujuanPeminjaman, gudangTujuan) {
  const action = 'ktf'
  let data = {}
  const permission = checkPermission(action)

  //proses kotak pinjam
  data.idKotakDipinjam = idKotakDipinjam.join(",")
  data.tujuanPeminjaman = tujuanPeminjaman
  data.gudangTujuan = gudangTujuan
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

function kotakTransfer(param) {
  var spreadsheet = SpreadsheetApp.openById(param.fileId)

  if (!spreadsheet) {
    return jsonResponse({
      'code': 404,
      'message': "File tidak ditemukan",
    })
  }

  var sheetPinjamKotak = spreadsheet.getSheetByName('Riwayat Transfer Antar Gudang')

  if (!sheetPinjamKotak) {
    return jsonResponse({
      'code': 404,
      'message': "Sheet tidak ditemukan",
    })
  }

  var lrPinjamKotak = sheetPinjamKotak.getLastRow()

  var sheetKotak = spreadsheet.getSheetByName('Kotak')

  if (!sheetKotak) {
    return jsonResponse({
      'code': 404,
      'message': "Sheet tidak ditemukan",
    })
  }

  var lrKotak = sheetKotak.getLastRow()

  var user = param.user
  var date = new Date()

  const idKotakDipinjam = param.idKotakDipinjam.split(",")

  idKotakDipinjam.forEach(id => {
    var dataRow = sheetKotak.getRange(2,1,lrKotak-1,11).getValues().findIndex((data)=>{return data[0]===id}) + 2
    sheetPinjamKotak.getRange(lrPinjamKotak+1,1).setValue(sheetKotak.getRange(dataRow,1).getValue())
    sheetPinjamKotak.getRange(lrPinjamKotak+1,2).setValue(sheetKotak.getRange(dataRow,5).getValue())
    sheetPinjamKotak.getRange(lrPinjamKotak+1,3).setValue(param.tujuanPeminjaman)
    sheetPinjamKotak.getRange(lrPinjamKotak+1,4).setValue(sheetKotak.getRange(dataRow,8).getValue())
    sheetPinjamKotak.getRange(lrPinjamKotak+1,5).setValue(param.gudangTujuan)
    sheetPinjamKotak.getRange(lrPinjamKotak+1,6).setValue(user)
    sheetPinjamKotak.getRange(lrPinjamKotak+1,7).setValue(Utilities.formatDate(date,"Asia/Bangkok","dd MMMM yyyy HH:mm:ss"))
    sheetKotak.getRange(dataRow,8).setValue(param.gudangTujuan)
    // sheetKotak.getRange(dataRow,12).setValue('Dipinjam')
    lrPinjamKotak = lrPinjamKotak + 1
  });

  return jsonResponse({
    'code': 200,
    'message': "Berhasil mencatat transfer kotak antar gudang",
  })
}