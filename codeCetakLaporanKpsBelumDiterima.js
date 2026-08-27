function showCetakLaporanKpsBelumDiterima() {
  var formCetakLaporanKpsBelumDiterima = HtmlService
    .createHtmlOutputFromFile('formCetakLaporanKpsBelumDiterima')
    .setWidth(400)
    .setHeight(200)
    
  SpreadsheetApp.getUi()
    .showModalDialog(formCetakLaporanKpsBelumDiterima, 'Form Cetak Laporan KPS Belum Diterima')
}

function clickBtnGenerateReportNotRecieve(formData) {
  const params = {
    periode: formData.periode
  }
  PropertiesService.getUserProperties().setProperty('paramTriggerReportNotReceive', JSON.stringify(params))

  ScriptApp.newTrigger('printReportKpsNotReceive')
    .timeBased()
    .after(5*1000)
    .create()

  return 'File masih diproses dan akan dikirim melalui email. Silahkan cek email secara berkala. Sesungguhnya Tuhan membersamai orang-orang yang sabar :)'
}

function printReportKpsNotReceive(e) {
  console.log('fungsi print Report Kps Not Receive ditrigger')

  const savedProperties = PropertiesService.getUserProperties().getProperty('paramTriggerReportNotReceive')

  var params

  if (savedProperties) {
    params = JSON.parse(savedProperties)
  } else {
    Logger.log('trigger tidak memiliki parameter')
    return
  }

  var periode = params.periode
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheetTemplate = spreadsheet.getSheetByName('Template Not Receive')
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
	sheetTemplate.getRange(4,3).setValue('KPS')
	sheetTemplate.getRange(4,4).setValue('Tanggal Awal')
  sheetTemplate.getRange(4,5).setValue('Tanggal Akhir')
  sheetTemplate.getRange(4,6).setValue('Keterangan')

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
		res[`${dataArea[i]}-Retail`] = defKps
		res[`${dataArea[i]}-Non Retail`] = defKps
		cr++
	}

	// mengambil kotak sesuai periode
	var filteredKotak = dataKotak.filter((kotak) => {
		var result = false
		kotak.periode.forEach(element => {
			var boxPeriode = Object.keys(element)[0]
      // console.log(kotak.labelKotak)
      // console.log(kotak.jenisDokumen)
      result = result || (parseInt(periode)===parseInt(boxPeriode) && kotak.status==='Terpakai' && kotak.jenisDokumen !== 'Consent Letter')
		});
		return result
	})
	// console.log('sebelum for each kotak yg difilter')
	// console.log(res)
	// menghapus kps yg terdapat dalam filtered kotak
	filteredKotak.forEach((data)=>{
		// console.log(data.labelKotak)
    // console.log(data.jenisDokumen)
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
				var excludedKps = res[`${boxArea}-${data.jenisDokumen}`].filter((el)=>!boxKps.includes(el))
        // console.log(excludedKps)
        res[`${boxArea}-${data.jenisDokumen}`] = excludedKps
			})
		})
	})

	// console.log('setelah for each kotak yg difilter')
  // console.log(res)

	cr = 5
	// res.forEach((data)=>{
	// 	sheetTemplate.getRange(cr,4).setValue(Object.values(data).join(', '))
	// 	cr++
	// })
  for (const key in res) {
    if (!Object.hasOwn(res, key)) continue;
    const [area, docType] = key.split('-')
    const element = res[key];
    element.forEach(kps => {
      const kpsDate = getKpsDate(kps, periode)
      sheetTemplate.getRange(cr,1).setValue(area)
      sheetTemplate.getRange(cr,2).setValue(docType)
      sheetTemplate.getRange(cr,3).setValue(kps)
      sheetTemplate.getRange(cr,4).setValue(kpsDate.startDate).setNumberFormat("dd-MMM-yyyy")
      sheetTemplate.getRange(cr,5).setValue(kpsDate.endDate).setNumberFormat("dd-MMM-yyyy")
      cr++
    });
  }

  
	// console.log(res)
  generateAndSendLink(`Laporan KPS Retail Dan Non Retail Yang Belum Diterima - ${periode}`, 'Template Not Receive')

  PropertiesService.getUserProperties().deleteProperty('paramTriggerReportNotReceive')

  console.log(e)
  if (e && e.triggerUid) {
    const triggerId = e.triggerUid
    deleteTriggerByUid(triggerId)
  }
  // return 'File selesai diproses dan akan dikirim melalui email. Silahkan cek email. Sesungguhnya Tuhan membersamai orang-orang yang sabar :)'
}