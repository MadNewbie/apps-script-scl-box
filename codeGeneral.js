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

function debug() {
  var docType = "Retail"
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = spreadsheet.getSheetByName('Master Jenis Dokumen')
  var lr = sheet.getLastRow()
  var masterDocTypeData = sheet.getRange(1,1,lr,2).getValues()
  var rowIndexSelected = masterDocTypeData.findIndex(x=>x[0]===docType)
  console.log(sheet.getRange(rowIndexSelected+1,2).getValue())
}