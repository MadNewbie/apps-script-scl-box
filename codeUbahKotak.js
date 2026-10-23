function showUbahKotakForm() {
  var formUbahKotak = HtmlService
    .createHtmlOutputFromFile('formUbahKotak')
    .setHeight(1750)
    .setWidth(2000)

  SpreadsheetApp.getUi()
    .showModalDialog(formUbahKotak, 'Form Ubah Kotak')
}

function getKodeKotak(pJenisKotak,pJenisDokumen,pGudang) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheetKotak = spreadsheet.getSheetByName('Kotak')
  var lrKotak = sheetKotak.getLastRow()
  var resFilteredKodeLabelKotak = Object.fromEntries(sheetKotak.getRange(2,1,lrKotak-1,9).getValues().filter(function(data){
    return data[1]===pJenisKotak && data[2]===pJenisDokumen && data[7]===pGudang && data[8] === 'Terpakai'
  }).map((data)=>[data[0],data[4]]))
  return resFilteredKodeLabelKotak
}

function getAreaPeriodeNoteOther(pKodeKotak) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheetKotak = spreadsheet.getSheetByName('Kotak')
  var lrKotak = sheetKotak.getLastRow()
  var dataRow = sheetKotak.getRange(2,1,lrKotak-1,11).getValues()
  var filteredData = dataRow.filter((data)=>{return data[0]===pKodeKotak})[0]
  var area = filteredData[5].split(", ")
  var note = filteredData[9]
  var otherDoc = filteredData[10]
  var periode = processedPeriode(filteredData[6], filteredData[2])
  var typeDoc = filteredData[2]
  var res = {typeDoc,area,note,otherDoc,periode}
  return res
}

function processedPeriode(data, typeDoc) {
  if (typeDoc === "Consent Letter") {
    return data.split(", ")
  } else {
    var res = []
    var splitDataPerPeriod = data.split(";")
    splitDataPerPeriod.forEach((data) => {
      var periode = data.match(/\((.*?)\)/)[1]
      var trimmedString = data.slice(4,-7)
      var periodeScnd = trimmedString.split(', ')
      var tmpPeriode = []
      periodeScnd.forEach((data) => {
        var periodeThrd = data.split("-")
        if(periodeThrd.length > 1){
          var first = parseInt(periodeThrd[0])
          var last = parseInt(periodeThrd[1])
          for (var i = first; i<=last;i++){
            tmpPeriode.push(i)
          }
        } else {
          tmpPeriode.push(periodeThrd)
        }
      })
      res.push({[periode]:tmpPeriode})
    })
    return res
  }
}

function updateDataKotak(formData) {
  const action = 'kuh'
  const data = {...formData}
  const permission = checkPermission(action)

  if (!permission.access) {
    return permission.message
  }

  //proses area
  data.area = formData.area.join(", ")

  //proses KPS
  let rawKps = []
  if (formData.jenisDokumen !== 'Consent Letter'){
    for(let i=formData.awalPeriode; i<=formData.akhirPeriode; i++){
      rawKps.push("KPS "+formatKps(formData[`kps[${i}]`])+" ("+ i +")")
    }
  }

  data.fileId = SpreadsheetApp.getActiveSpreadsheet().getId()
  data.user = permission.user
  data.action = action
  data.rawKps = rawKps.join(";")

  
  const responseFromBE = sendDataToBE(data)

  if (responseFromBE.success) {
    return responseFromBE.message
  } else {
    throw new Error(responseFromBE.message)
  }
}

function kotakUbah(param) {
  var spreadsheet = SpreadsheetApp.openById(param.fileId)

  if (!spreadsheet) {
    return jsonResponse({
      'code': 404,
      'message': "File tidak ditemukan",
    })
  }

  var sheet = spreadsheet.getSheetByName('Kotak')
  
  if (!sheet) {
    return jsonResponse({
      'code': 404,
      'message': "Sheet tidak ditemukan",
    })
  }

  var lrKotak = sheet.getLastRow()
  var dataRow = sheet.getRange(2,1,lrKotak-1,1).getValues().findIndex((data)=>{return data[0]===param.kodeKotak}) + 2
  var user = param.user
  var date = new Date()

  sheet.getRange(dataRow, 2).setValue(param.jenisKotak)
  sheet.getRange(dataRow, 3).setValue(param.jenisDokumen)
  sheet.getRange(dataRow, 5).setValue(param.labelKotak)
  sheet.getRange(dataRow, 6).setValue(param.area)
  if (param.jenisDokumen != 'Consent Letter') {
    sheet.getRange(dataRow, 7).setValue(param.rawKps)
  }else{
    sheet.getRange(dataRow, 7).setValue(param.periode)
  }
  sheet.getRange(dataRow, 8).setValue(param.gudang)
  sheet.getRange(dataRow, 10).setValue(param.catatan)
  sheet.getRange(dataRow, 11).setValue(param.dokumenLain)
  sheet.getRange(dataRow, 16).setValue(Utilities.formatDate(date,"Asia/Bangkok","dd MMMM yyyy HH:mm:ss"))
  sheet.getRange(dataRow, 17).setValue(user)

  return jsonResponse({
    'code':200,
    'message':'Berhasil mengubah data'
  })
}

// function debug() {
//   var pKodeKotak = 'SCL.2'
//   var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
//   var sheetKotak = spreadsheet.getSheetByName('Kotak')
//   var lrKotak = sheetKotak.getLastRow()
//   var dataRow = sheetKotak.getRange(2,1,lrKotak-1,11).getValues()
//   var filteredData = dataRow.filter((data)=>{return data[0]===pKodeKotak})[0]
//   var area = filteredData[5].split(", ")
//   var note = filteredData[9]
//   var otherDoc = filteredData[10]
//   var periode = processedPeriode(filteredData[6], filteredData[2])
//   var res = {area,note,otherDoc,periode}
// }