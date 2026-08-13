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

// function debug() {
//   var docType = "Retail"
//   var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
//   var sheet = spreadsheet.getSheetByName('Master Jenis Dokumen')
//   var lr = sheet.getLastRow()
//   var masterDocTypeData = sheet.getRange(1,1,lr,2).getValues()
//   var rowIndexSelected = masterDocTypeData.findIndex(x=>x[0]===docType)
//   console.log(sheet.getRange(rowIndexSelected+1,2).getValue())
// }