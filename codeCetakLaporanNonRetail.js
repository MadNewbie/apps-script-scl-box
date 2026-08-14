function showCetakLaporanNonRetail() {
  var formCetakLaporanNonRetail = HtmlService
    .createHtmlOutputFromFile('formCetakLaporanNonRetail')
    .setWidth(400)
    .setHeight(200)
    
  SpreadsheetApp.getUi()
    .showModalDialog(formCetakLaporanNonRetail, 'Form Cetak Laporan Non Retail')
}

function printReportNonRetail(formData) {
    var periode = formData.periode
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
    var sheetTemplate = spreadsheet.getSheetByName('Report Template')
    var sheetKotak = spreadsheet.getSheetByName('Kotak')
    var sheetArea = spreadsheet.getSheetByName('Master Area')
    var lrKotak = sheetKotak.getLastRow()
    var lrArea = sheetKotak.getLastRow()
    var date = new Date()
    sheetTemplate.clear()

    var dataArea = sheetArea.getRange(1,2,lrArea,1).getValues()
    var dataKotak = getAllDataAsObject()

    //generate header
    sheetTemplate.getRange(1,1).setValue('Laporan KPS Non Retail')
    sheetTemplate.getRange(2,1).setValue('Periode')
    sheetTemplate.getRange(2,2).setValue(periode)
    sheetTemplate.getRange(3,1).setValue('Tanggal Cetak')
    sheetTemplate.getRange(3,2).setValue(Utilities.formatDate(date,"Asia/Bangkok","dd MMMM yyyy HH:mm:ss"))
    sheetTemplate.getRange("A4:A5").setValue('Area').merge().setHorizontalAlignment("center").setVerticalAlignment("center")
    sheetTemplate.getRange("B4:BA4").setValue('KPS').merge().setHorizontalAlignment("center")
    sheetTemplate.getRange("BB4:BC4").setValue('Jumlah Data').merge().setHorizontalAlignment("center")
    for(var i=1; i<=52; i++){
        sheetTemplate.getRange(5,i+1).setValue(i).setHorizontalAlignment("center")
    }
    sheetTemplate.getRange(5,54).setValue('Sudah Diterima').setHorizontalAlignment("center")
    sheetTemplate.getRange(5,55).setValue('Belum Diterima').setHorizontalAlignment("center")
    cr = 5
    for(var i=0; i<dataArea.length; i++){
        sheetTemplate.getRange(cr+i+1,1).setValue(dataArea[i])
    }

    var filteredKotak = dataKotak.filter((kotak) => {
        var res = false
        kotak.periode.forEach((data) => {
            var boxPeriode = Object.keys(data)[0]
            res = res || (parseInt(periode)===parseInt(boxPeriode) && kotak.jenisDokumen==='Non Retail' && kotak.status==='Terpakai')
        })
        return res
    })

    // console.log(filteredKotak)

    // memberi centang pada cell yang sudah ada

    filteredKotak.forEach((data) => {
        var lrTemplate = sheetTemplate.getLastRow()
        var dataRef = sheetTemplate.getRange(6,1,lrTemplate-5,1).getValues().map((data)=>data[0])
        // console.log(dataRef)
        data.area.forEach((area) => {
            var dataRow = 6 + dataRef.indexOf(area)
            data.periode.filter((dataInside) => {
                var boxPeriode = Object.keys(dataInside)[0]
                return parseInt(periode) === parseInt(boxPeriode)
            }).forEach((dataInside) => {
                var boxKps = Object.values(dataInside)[0]
                boxKps.forEach((kps) => {
                    var dataCol = 1 + kps
                    // console.log('Di dalam foreach kps:')
                    // console.log(dataRow)
                    // console.log(dataCol)
                    sheetTemplate.getRange(dataRow, dataCol).setValue('V')
                })
            })
        })
    })
    
    for(var i=0; i<dataArea.length; i++){
        sheetTemplate.getRange(cr+i+1,54).setValue(`=countif(B${cr+i+1}:BA${cr+i+1},"V")`)
        sheetTemplate.getRange(cr+i+1,55).setValue(`=countif(B${cr+i+1}:BA${cr+i+1},"")`)
    }

    var dataFile = generateAndSendLink(`Laporan KPS Non Retail Periode - ${periode}`)
    return 'File selesai diproses dan akan dikirim melalui email. Silahkan cek email. Sesungguhnya Tuhan membersamai orang-orang yang sabar :)'
}