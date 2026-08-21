function getDropdownJenisKotakData() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = spreadsheet.getSheetByName('Master Jenis Kotak')

  var range = sheet.getRange('A:A')
  var values = range.getValues()

  var data = values.filter(function(row) {
    return row[0] !== '' && row[0] !== undefined
  })

  var dropdownData = data.map(function(row) {
    return row[0]
  })

  return dropdownData
}

function getDropdownGudangData() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = spreadsheet.getSheetByName('Master Gudang')

  var range = sheet.getRange('A:A')
  var values = range.getValues()

  var data = values.filter(function(row) {
    return row[0] !== '' && row[0] !== undefined
  })

  var dropdownData = data.map(function(row) {
    return row[0]
  })

  return dropdownData
}

function getDrowpdownJenisDokumenByJenisKotakData(pBoxType) {
  var boxType = pBoxType
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = spreadsheet.getSheetByName('Mapping Jenis Kotak Jenis Dokumen')

  var range = sheet.getRange('A:B')
  var values = range.getValues()

  var data = values.filter(function(row) {
    if (row[1] === boxType){
      return row[0] 
    }
  })

  var dropdownData = data.map(function(row) {
    return row[0]
  })

  return dropdownData
}

function getDropdownJenisDokumenData() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = spreadsheet.getSheetByName('Master Jenis Dokumen')

  var range = sheet.getRange('A:A')
  var values = range.getValues()

  var data = values.filter(function(row) {
    return row[0] !== '' && row[0] !== undefined
  })

  var dropdownData = data.map(function(row) {
    return row[0]
  })

  return dropdownData
}

function getDropdownAreaData() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = spreadsheet.getSheetByName('Master Area')

  var range = sheet.getRange('B:B')
  var values = range.getValues()

  var data = values.filter(function(row) {
    return row[0] !== '' && row[0] !== undefined
  })

  var dropdownData = data.map(function(row) {
    return row[0]
  })

  return dropdownData
}

function getBoxLabel(pBoxType,pDocType,pYear) {
  var boxType = pBoxType
  var docType = pDocType
  var year = pYear
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = spreadsheet.getSheetByName('Kotak')
  var sheetDocType = spreadsheet.getSheetByName('Master Jenis Dokumen')
  var lastRow = sheet.getLastRow()
  var lrDocType = sheetDocType.getLastRow()

  //mengambil nomor berdasarkan jenis kotak per tahun
  nextNoYear = lastRow == 1 ? 1 : sheet.getRange(2,2,lastRow-1,13).getValues().filter(x=>(x[0]===boxType)&&(x[1]===docType)&&(new Date(x[12]).getFullYear()===year)).length + 1

  //mengambil kode jenis dokumen
  docTypeCode = sheetDocType.getRange(sheetDocType.getRange(1,1,lrDocType,2).getValues().findIndex(x=>x[0]===docType)+1,2).getValue()

  return `${boxType}/${docTypeCode}/${year}/${nextNoYear}`
}

function generateFileAndGetUrl(pNameFile){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = spreadsheet.getSheetByName('Report Template')
  var newFile = SpreadsheetApp.create(pNameFile)
  var newFileId = newFile.getId()

  sheet.copyTo(newFile)

  SpreadsheetApp.flush();

  Utilities.sleep(1000);

  var defSheet = newFile.getSheetByName('Sheet1')
  if (defSheet) {
    newFile.deleteSheet(defSheet)
  }

  var downloadUrl = `https://docs.google.com/spreadsheets/d/${newFileId}/export?format=xlsx&id=${newFileId}`

  return {
    downloadUrl: downloadUrl,
    newFileId: newFileId
  }
}

function cleanupFileAndSheet(data) {
  try {
    DriveApp.getFileById(data.newFileId).setTrashed(true)
  } catch (error) {
    Logger.log(`Error while deleting drive file: ${error.message}`)
  }

  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
    var sheet = spreadsheet.getSheetByName('Report Template')
    if (sheet && spreadsheet.getSheets().length > 1){
      sheet.clear()
    }
  } catch (error) {
    Logger.log(`Error while clearing up sheet: ${error.message}`)
  }
}

function generateAndCleanupExcel(pNamefile) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheetToExport = spreadsheet.getSheetByName('Report Template')
  var newFile = SpreadsheetApp.create(pNamefile)
  var newFileId = newFile.getId()

  sheetToExport.copyTo(newFile)

  SpreadsheetApp.flush()

  Utilities.sleep(1000)

  var url = `https://docs.google.com/spreadsheets/d/${newFileId}/export?format=xlsx`
  var token = ScriptApp.getOAuthToken()
  var response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': `Bearer ${token}`
    }
  })


  var folder = DriveApp.getFileById(spreadsheet.getId()).getParents().next()
  var blob = response.getBlob().setName(`${pNamefile}.xlsx`)
  var file = folder.createFile(blob)

  DriveApp.getFileById(newFile.getId()).setTrashed(true)

  sheetToExport.clear()

  return {
    downloadUrl: file.getUrl()
  }
}

function generateAndSendLink(pNameFile) {
  // generate file
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = spreadsheet.getSheetByName('Report Template')
  var newFile = SpreadsheetApp.create(pNameFile)
  var newFileId = newFile.getId()
  
  const tresholdFile = 15*60*1000
  
  var copiedSheet = sheet.copyTo(newFile)

  copiedSheet.setName(pNameFile)
  copiedSheet.showSheet()

  var defSheet = newFile.getSheetByName('Sheet1')
  if (defSheet) {
    newFile.deleteSheet(defSheet)
  }

  ScriptApp.newTrigger('deleteSentFile')
    .timeBased()
    .after(tresholdFile)
    .create()
  
  PropertiesService.getScriptProperties().setProperties({'FILE_TO_DELETE': newFileId})

  SpreadsheetApp.flush()

  //sent email
  var email = Session.getActiveUser().getEmail()

  try {
    var newFileUrl = newFile.getUrl()

    var subject = `Your Google Drive Link: ${pNameFile}`
    var htmlBody = `
      <p>Halo,</p>
      <p>Anda telah meminta untuk membuatkan laporan dan berikut adalah tautan laporan yang anda minta:</p><br>
      <p><a href="${newFileUrl}">${pNameFile}</a></p>
      <br>
      <p>Harap segera unduh file tersebut, karena sistem hanya menyimpannya dalam waktu kurang lebih 15 menit dari email ini diterima.</p>
      <br>
      <br>
      <p>Semoga hari anda menyenangkan,<br>Sistem Otomatis</p>
    `

    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: htmlBody
    })

    console.log(`Email berhasil terkirim ke ${user}`)
  } catch (error) {
    console.log(`Error terjadi ketika menjalankan script: ${error.toString()}`)
  }
}

function deleteSentFile() {
  const scriptProperties = PropertiesService.getScriptProperties()
  const fileId = scriptProperties.getProperty('FILE_TO_DELETE')

  if (fileId) {
    try {
      DriveApp.getFileById(fileId).setTrashed(true)
      Logger.log(`Successfully moved file ${fileId} to trash`)
    } catch (error) {
      Logger.log(`Error while deleting file: ${error.message}`)
    }  finally {
      scriptProperties.deleteProperty('FILE_TO_DELETE')
    }
  }
}

function getAllDataAsObject() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheetKotak = spreadsheet.getSheetByName('Kotak')
  var lrKotak = sheetKotak.getLastRow()
  var res = []

  for (var i = 2; i<=lrKotak; i++) {
    var modelBox = {}
    if (sheetKotak.getRange(i,3).getValue()!=='Consent Letter') {
      modelBox['periode'] = getPeriodeKpsDataAsObject(sheetKotak.getRange(i,7).getValue())
    } else {
      modelBox['periode'] = sheetKotak.getRange(i,7).getValue().split(', ')
    }
    modelBox['kodeKotak'] = sheetKotak.getRange(i,1).getValue()
    modelBox['jenisKotak'] = sheetKotak.getRange(i,2).getValue()
    modelBox['jenisDokumen'] = sheetKotak.getRange(i,3).getValue()
    modelBox['nomorKotak'] = sheetKotak.getRange(i,4).getValue()
    modelBox['labelKotak'] = sheetKotak.getRange(i,5).getValue()
    modelBox['area'] = sheetKotak.getRange(i,6).getValue().split(', ')
    modelBox['lokasi'] = sheetKotak.getRange(i,8).getValue()
    modelBox['status'] = sheetKotak.getRange(i,9).getValue()
    modelBox['catatan'] = sheetKotak.getRange(i,10).getValue()
    modelBox['dokumenLain'] = sheetKotak.getRange(i,11).getValue()
    modelBox['statusKhusus'] = sheetKotak.getRange(i,12).getValue()
    res.push(modelBox)
  }

  return res
}

function getPeriodeKpsDataAsObject(stringData) {
  var res = []
  var firstFormattedData = stringData.split(';')
  firstFormattedData.forEach((data) => {
    var indexOfKurBuk = data.indexOf('(')
    var indexOfKurTup = data.indexOf(')')
    var periode = parseInt(data.substring(indexOfKurBuk+1, indexOfKurTup))
    var kpsString = data.substring(4, indexOfKurBuk - 1)
    var kpsStringScnd = kpsString.split(', ')
    var dataKpsArr = []
    kpsStringScnd.forEach((insideData) => {
      var kpsStringThrd = insideData.split('-')
      if(kpsStringThrd.length > 1){
        var first = parseInt(kpsStringThrd[0])
        var last = parseInt(kpsStringThrd[1])
        for (var i=first; i<=last; i++) {
          dataKpsArr.push(i)
        }
      } else {
        dataKpsArr.push(parseInt(kpsStringThrd[0]))
      }
    })
    res.push({[periode]: dataKpsArr})
  })
  return res
}

function getKpsDate(kps, periode) {
  if (periode > 1000) {
    firstDayOfYear = new Date(periode, 0, 1)
    firstDayOfWeek = firstDayOfYear.getDay()
    offset = firstDayOfWeek === 0 ? 0 : firstDayOfWeek

    startDate = new Date(
      firstDayOfYear.getFullYear(),
      0,
      1-offset+1+(kps-1)*7
    )

    endDate = new Date(startDate)
    endDate.setDate(startDate.getDate() + 6)
  }

  return {
    startDate: startDate,
    endDate: endDate
  }
}

// function debug() {
//   var kotak = getAllDataAsObject()
//   console.log(kotak)
// }