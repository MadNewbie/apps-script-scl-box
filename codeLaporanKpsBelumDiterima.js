function showCetakLaporanKpsBelumDiterima() {
  var formCetakLaporanKpsBelumDiterima = HtmlService
    .createHtmlOutputFromFile('formCetakLaporanKpsBelumDiterima')
    .setWidth(400)
    .setHeight(200)
    
  SpreadsheetApp.getUi()
    .showModalDialog(formCetakLaporanKpsBelumDiterima, 'Form Cetak Laporan KPS Belum Diterima')
}

function printReportKpsNotReceive(formData) {
  var periode = formData.periode
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheetTemplate = spreadsheet.getSheetByName('Report Template')
  var sheetKotak = spreadsheet.getSheetByName('Kotak')
  var sheetArea = spreadsheet.getSheetByName('Master Area')
  var lrKotak = sheetKotak.getLastRow()
  var lrArea = sheetKotak.getLastRow()
  var date = new Date()
  var res = {}
  sheetTemplate.clear()

	//generate header
	sheetTemplate.getRange(1,1).setValue('Laporan KPS Yang Belum Diterima')
  sheetTemplate.getRange(2,1).setValue('Periode')
  sheetTemplate.getRange(2,2).setValue(periode)
  sheetTemplate.getRange(3,1).setValue('Tanggal Cetak')
  sheetTemplate.getRange(3,2).setValue(Utilities.formatDate(date,"Asia/Bangkok","dd MMMM yyyy HH:mm:ss"))
	sheetTemplate.getRange(4,1).setValue('Area')
	sheetTemplate.getRange(4,2).setValue('Jenis Dokumen')
	sheetTemplate.getRange(4,3).setValue('Periode')
	sheetTemplate.getRange(4,4).setValue('KPS')

	cr = 4

	var dataKotak = getAllDataAsObject()
	var dataArea = sheetArea.getRange(1,2,lrArea,1).getValues().map((data)=>data[0]).filter((data)=>{return data!==""})
	// console.log(dataArea)
	var defKps = []
	for(var c=1; c<=52; c++){
		defKps.push(c)		
	}

	//membuat model default data
	for(var i=0;i<dataArea.length;i++){
		sheetTemplate.getRange(cr+i+1,1).setValue(dataArea[i])
		sheetTemplate.getRange(cr+i+2,1).setValue(dataArea[i])
		sheetTemplate.getRange(cr+i+1,2).setValue('Retail')
		sheetTemplate.getRange(cr+i+2,2).setValue('Non Retail')
		sheetTemplate.getRange(cr+i+1,3).setValue(periode)
		sheetTemplate.getRange(cr+i+2,3).setValue(periode)
		res[`${dataArea[i]} Retail`] = defKps
		res[`${dataArea[i]} Non Retail`] = defKps
		cr++
	}

	// mengambil kotak sesuai periode
	var filteredKotak = dataKotak.filter((kotak) => {
		var result = false
		kotak.periode.forEach(element => {
			var boxPeriode = Object.keys(element)[0]
      		result = result || (parseInt(periode)===parseInt(boxPeriode) && kotak.status==='Terpakai')
		});
		return result
	})
	// console.log('sebelum for each kotak yg difilter')
	// console.log(res)
	// menghapus kps yg terdapat dalam filtered kotak
	filteredKotak.forEach((data)=>{
		// console.log(data)
		data.area.forEach((area) => {
			var boxArea = area
			data.periode.filter((dataInside) => {
        // console.log('di dalam filter periode')
				var boxPeriode = Object.keys(dataInside)[0]
				return parseInt(periode) === parseInt(boxPeriode)
			}).forEach((dataInside) => {
				var boxKps = Object.values(dataInside)[0]
        // console.log(boxKps)
        // console.log(res[`${boxArea} ${data.jenisDokumen}`])
        // console.log(res[resKey])
				var excludedKps = res[`${boxArea} ${data.jenisDokumen}`].filter((el)=>!boxKps.includes(el))
        // console.log(excludedKps)
        res[`${boxArea} ${data.jenisDokumen}`] = excludedKps
			})
		})
	})

	console.log('setelah for each kotak yg difilter')
  console.log(res)

	cr = 5
	// res.forEach((data)=>{
	// 	sheetTemplate.getRange(cr,4).setValue(Object.values(data).join(', '))
	// 	cr++
	// })
  for (const key in res) {
    if (!Object.hasOwn(res, key)) continue;
    
    const element = res[key];
    sheetTemplate.getRange(cr,4).setValue(element.join(', '))
		cr++
    
  }

	// console.log(res)
  // generateAndSendLink(`Laporan KPS Retail Dan Non Retail Yang Belum Diterima - ${periode}`)
  return 'File selesai diproses dan akan dikirim melalui email. Silahkan cek email. Sesungguhnya Tuhan membersamai orang-orang yang sabar :)'
}