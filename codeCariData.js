function showCariData() {
  var formCariData = HtmlService
    .createHtmlOutputFromFile('formCariData')
    .setWidth(1000)
    .setHeight(750)
    
  SpreadsheetApp.getUi()
    .showModalDialog(formCariData, 'Form Cari Data')
}

function getSearchDataByParam(pArea, pJenisDokumen, pKps, pPeriode) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheetKotak = spreadsheet.getSheetByName('Kotak')
  var lrKotak = sheetKotak.getLastRow()
  var dataRow = sheetKotak.getRange(2,1,lrKotak-1,11).getValues()
  var firstFilteredData = dataRow.filter((item) => { return((pArea==='' || pArea === item[5]) && (pJenisDokumen==='' || pJenisDokumen === item[2]) && item[8]!=='Rusak') })
  // console.log(firstFilteredData)
  var res = firstFilteredData.filter((item) => {
    var rawKps = item[6]
    var periodeAndKps = getPeriodesAndKpses(item[6])
    var result = false
    // console.log(periodeAndKps)
    periodeAndKps.forEach((item) => {
      var periode = Object.keys(item)[0]
      var data = Object.values(item)[0]
      // console.log(periode)
      // console.log(data)
      // console.log(`pPeriode==='': ${pPeriode===''}`)
      // console.log(`parseInt(pPeriode)===parseInt(periode): ${parseInt(pPeriode)===parseInt(periode)}`)
      // console.log(`pKps==='': ${pKps===''}`)
      // console.log(`data.includes(pKps): ${data.includes(parseInt(pKps))}`)
      result = result || ((pPeriode===''||parseInt(pPeriode)===parseInt(periode)) && (pKps===''||data.includes(parseInt(pKps))))
    })
    // console.log(result)
    return result
  })
  // console.log(res)
  return res
}

function getPeriodesAndKpses(data) {
  var res = []
  var firstFormattedData = data.split(';')
  firstFormattedData.forEach((item) => {
    var indexOfKurBuk = item.indexOf('(')
    var indexOfKurTup = item.indexOf(')')
    var periode = item.substring(indexOfKurBuk + 1, indexOfKurTup)
    var dataKpsString = item.substring(4, indexOfKurBuk-1)
    var dataKpsStringScnd = dataKpsString.split(', ')
    var dataKpsArr = []
    dataKpsStringScnd.forEach((insideItem) => {
      var dataKpsStringThrd = insideItem.split('-')
      if (dataKpsStringThrd.length > 1) {
        var first = parseInt(dataKpsStringThrd[0])
        var last = parseInt(dataKpsStringThrd[1])
        for (var i=first; i<=last; i++) {
          dataKpsArr.push(i)
        }
      } else {
        dataKpsArr.push(dataKpsStringThrd)
      }
    })
    res.push({[periode]:dataKpsArr})
  })
  return res
}

// function debug() {
//   getSearchDataByParam("","","3","2025")
// }