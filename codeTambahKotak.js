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
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = spreadsheet.getSheetByName('Kotak')
  var jmlKotakSheet = spreadsheet.getSheetByName('Jumlah Kotak')
  var lastRow = sheet.getLastRow()
  var lrJmlKotak = jmlKotakSheet.getLastRow()
  var user = Session.getActiveUser().getEmail()

  //proses area
  area = formData.area.join(", ")

  //proses KPS
  if (formData.jenisDokumen !== 'Consent Letter'){
    rawKps = []
    for(let i=formData.awalPeriode; i<=formData.akhirPeriode; i++){
      rawKps.push("KPS "+formatKps(formData[`kps[${i}]`])+" ("+ i +")")
    }
  }

  //nomor kotak selanjutnya
  rowJenisKotak = jmlKotakSheet.getRange(1,1,lrJmlKotak,1).getValues().findIndex(row=>row[0]===formData.jenisKotak)+1
  nextNo = jmlKotakSheet.getRange(rowJenisKotak,2).getValue() - jmlKotakSheet.getRange(rowJenisKotak,6).getValue() + 1

  date = new Date()
  year = new Date().getFullYear()

  //nomor kotak per tahun per jenis
  nextNoYear = lastRow == 1 ? 1 : sheet.getRange(2,2,lastRow-1,13).getValues().filter(x=>(x[0]===formData.jenisKotak)&&(x[1]===formData.jenisDokumen)&&(new Date(x[12]).getFullYear()===year)).length + 1

  sheet.getRange(lastRow + 1, 1).setValue(formData.jenisKotak+"."+nextNo)
  sheet.getRange(lastRow + 1, 2).setValue(formData.jenisKotak)
  sheet.getRange(lastRow + 1, 3).setValue(formData.jenisDokumen)
  sheet.getRange(lastRow + 1, 4).setValue(nextNoYear)
  sheet.getRange(lastRow + 1, 5).setValue(formData.labelKotak)
  sheet.getRange(lastRow + 1, 6).setValue(area)
  if (formData.jenisDokumen != 'Consent Letter') {
    sheet.getRange(lastRow + 1, 7).setValue(rawKps.join(";"))
  }else{
    sheet.getRange(lastRow + 1, 7).setValue(formData.periode)
  }
  sheet.getRange(lastRow + 1, 8).setValue(formData.gudang)
  sheet.getRange(lastRow + 1, 9).setValue('Terpakai')
  sheet.getRange(lastRow + 1, 10).setValue(formData.catatan)
  sheet.getRange(lastRow + 1, 11).setValue(formData.dokumenLain)
  sheet.getRange(lastRow + 1, 14).setValue(Utilities.formatDate(date,"Asia/Bangkok","dd MMMM yyyy HH:mm:ss"))
  sheet.getRange(lastRow + 1, 15).setValue(user)

  return 'Berhasil menambahkan data'
}