function showPengembalianKotakAntarGudangSementara() {
  var formPengembalianKotakAntarGudangSementara = HtmlService
    .createHtmlOutputFromFile('formPengembalianKotakAntarGudangSementara')
    .setWidth(1000)
    .setHeight(750)

    SpreadsheetApp.getUi()
      .showModalDialog(formPengembalianKotakAntarGudangSementara, 'Form Pengembalian Kotak Antar Gudang Sementara')
}

function getAvailDataKotakForPengembalianSementara(gudangAsal) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheetRiwayatPinjamSementara = spreadsheet.getSheetByName('Riwayat Pinjam Kotak Antar Gudang Sementara')
  var lrRiwayatPinjamSementara = sheetRiwayatPinjamSementara.getLastRow()
  var user = Session.getActiveUser().getEmail()

  var dataKotakDipinjam = sheetRiwayatPinjamSementara.getRange(2,1,lrRiwayatPinjamSementara-1,7).getValues()
  var res = dataKotakDipinjam.filter(data => {return data[4]===gudangAsal && !data[5] && data[6]===user})
  return res
}

function saveDataPengembalianKotakAntarGudangSementara(idKotak, gudangTujuan) {
  const action = 'kpk'
  let data = {}
  const permission = checkPermission(action)

  //proses kotak pinjam
  data.idKotak = idKotak.join(",")
  data.gudangTujuan = gudangTujuan
  data.fileId = SpreadsheetApp.getActiveSpreadsheet().getId()
  data.user = permission.user
  data.action = action

  console.log(data)

  const responseFromBE = sendDataToBE(data)

  if (responseFromBE.success) {
    return responseFromBE.message
  } else {
    throw new Error(responseFromBE.message)
  }
}

function kotakPengembalianKotak(param) {
  var spreadsheet = SpreadsheetApp.openById(param.fileId)

  if (!spreadsheet) {
    return jsonResponse({
      'code': 404,
      'message': "File tidak ditemukan",
    })
  }

  var sheetPinjamKotak = spreadsheet.getSheetByName('Riwayat Pinjam Kotak Antar Gudang Sementara')

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

  const idKotak = param.idKotak.split(",")

  idKotak.forEach(id => {
    var dataRow = sheetKotak.getRange(2,1,lrKotak-1,11).getValues().findIndex((data)=>{return data[0]===id}) + 2
    var dataPinjamRow = sheetPinjamKotak.getRange(2,1,lrPinjamKotak-1,1).getValues().findIndex((data)=>{return data[0]===id})+2
    sheetKotak.getRange(dataRow,8).setValue(param.gudangTujuan)
    sheetKotak.getRange(dataRow,12).setValue('')
    sheetPinjamKotak.getRange(dataPinjamRow,6).setValue('TRUE')
    sheetPinjamKotak.getRange(dataPinjamRow,9).setValue(Utilities.formatDate(date,"Asia/Bangkok","dd MMMM yyyy HH:mm:ss"))
    sheetPinjamKotak.getRange(dataPinjamRow,10).setValue(param.gudangTujuan)
  });

  return jsonResponse({
    'code': 200,
    'message': "Berhasil mencatat pengembalian kotak",
  })
}