function showTambahKotakForm() {
  var formTambahKotak = HtmlService
    .createHtmlOutputFromFile('formTambahKotak')
    .setHeight(1750)
    .setWidth(2000)

  SpreadsheetApp.getUi()
    .showModalDialog(formTambahKotak, 'Form Tambah Kotak')
}

function formatKps(data){
  // console.log(JSON.stringify(data))
  // var arrayData = data.split(",")
  data = data.map((data)=>parseInt(data))
  var kpsWithDash=[]
  var x=0
  while(x<=data.length){
    var y=x+1
    if(x<=data.length-1){
      if(data[y]!=data[x]+1){
        kpsWithDash.push(`${data[x]}`)
        x=y
        continue
      }
    }else{
      break
    }
    while (y<=data.length) {
      if(data[y+1]!=data[y]+1){
        kpsWithDash.push(`${data[x]} - ${data[y]}`)
        if(y+1<=data.length){
          x=y+1
        }
        break
      }
      y+=1
    }
  }
  res = kpsWithDash.join(", ")
  return res
}

function saveDataTambahKotak(formData) {
  const action = 'kth'
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

function kotakTambah(param) {
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

  var jmlKotakSheet = spreadsheet.getSheetByName('Jumlah Kotak')

  if (!jmlKotakSheet) {
    return jsonResponse({
      'code': 404,
      'message': "Sheet tidak ditemukan",
    })
  }

  var lastRow = sheet.getLastRow()
  var lrJmlKotak = jmlKotakSheet.getLastRow()
  var user = param.user

  //nomor kotak selanjutnya
  rowJenisKotak = jmlKotakSheet.getRange(1,1,lrJmlKotak,1).getValues().findIndex(row=>row[0]===param.jenisKotak)+1
  nextNo = jmlKotakSheet.getRange(rowJenisKotak,2).getValue() - jmlKotakSheet.getRange(rowJenisKotak,6).getValue() + 1

  date = new Date()
  year = new Date().getFullYear()

  //nomor kotak per tahun per jenis
  nextNoYear = lastRow == 1 ? 1 : sheet.getRange(2,2,lastRow-1,13).getValues().filter(x=>(x[0]===param.jenisKotak)&&(x[1]===param.jenisDokumen)&&(new Date(x[12]).getFullYear()===year)&&(x[7]==='Terpakai')).length + 1

  sheet.getRange(lastRow + 1, 1).setValue(param.jenisKotak+"."+nextNo)
  sheet.getRange(lastRow + 1, 2).setValue(param.jenisKotak)
  sheet.getRange(lastRow + 1, 3).setValue(param.jenisDokumen)
  sheet.getRange(lastRow + 1, 4).setValue(nextNoYear)
  sheet.getRange(lastRow + 1, 5).setValue(param.labelKotak)
  sheet.getRange(lastRow + 1, 6).setValue(param.area)
  if (param.jenisDokumen != 'Consent Letter') {
    sheet.getRange(lastRow + 1, 7).setValue(param.rawKps)
  }else{
    sheet.getRange(lastRow + 1, 7).setValue(param.periode)
  }
  sheet.getRange(lastRow + 1, 8).setValue(param.gudang)
  sheet.getRange(lastRow + 1, 9).setValue('Terpakai')
  sheet.getRange(lastRow + 1, 10).setValue(param.catatan)
  sheet.getRange(lastRow + 1, 11).setValue(param.dokumenLain)
  sheet.getRange(lastRow + 1, 14).setValue(Utilities.formatDate(date,"Asia/Bangkok","dd MMMM yyyy HH:mm:ss"))
  sheet.getRange(lastRow + 1, 15).setValue(user)

  return jsonResponse({
    'code':200,
    'message':'Berhasil menambahkan data'
  })
}