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
  var resFilteredKodeLabelKotak = Object.fromEntries(sheetKotak.getRange(2,1,lrKotak-1,8).getValues().filter(function(data){
    return data[1]===pJenisKotak && data[2]===pJenisDokumen && data[7]===pGudang
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
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheet =  spreadsheet.getSheetByName('Kotak')
  var lrKotak = sheet.getLastRow()
  var dataRow = sheet.getRange(2,1,lrKotak-1,1).getValues().findIndex((data)=>{return data[0]===formData.kodeKotak}) + 2
  var user = Session.getActiveUser().getEmail()
  var date = new Date()

  //proses area
  area = formData.area.join(", ")

  //proses KPS
  if (formData.jenisDokumen !== 'Consent Letter'){
    rawKps = []
    for(let i=formData.awalPeriode; i<=formData.akhirPeriode; i++){
      rawKps.push("KPS "+formatKps(formData[`kps[${i}]`])+" ("+ i +")")
    }
  }

  sheet.getRange(dataRow, 2).setValue(formData.jenisKotak)
  sheet.getRange(dataRow, 3).setValue(formData.jenisDokumen)
  sheet.getRange(dataRow, 5).setValue(formData.labelKotak)
  sheet.getRange(dataRow, 6).setValue(area)
  if (formData.jenisDokumen != 'Consent Letter') {
    sheet.getRange(dataRow, 7).setValue(rawKps.join(";"))
  }else{
    sheet.getRange(dataRow, 7).setValue(formData.periode)
  }
  sheet.getRange(dataRow, 8).setValue(formData.gudang)
  sheet.getRange(dataRow, 10).setValue(formData.catatan)
  sheet.getRange(dataRow, 11).setValue(formData.dokumenLain)
  sheet.getRange(dataRow, 16).setValue(Utilities.formatDate(date,"Asia/Bangkok","dd MMMM yyyy HH:mm:ss"))
  sheet.getRange(dataRow, 17).setValue(user)
  return 'Berhasil mengubah data'
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