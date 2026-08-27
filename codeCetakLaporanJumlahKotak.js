function clickBtnGenerateReportAmount() {
  ScriptApp.newTrigger('printReportAmountNumber')
    .timeBased()
    .after(5*1000)
    .create()

  window.alert('File masih diproses dan akan dikirim melalui email. Silahkan cek email secara berkala. Sesungguhnya Tuhan membersamai orang-orang yang sabar :)')
}

function printReportAmountNumber(e) {
  console.log('fungsi print Report Jumlah Kotak ditrigger')

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheetTemplate = spreadsheet.getSheetByName('Template Box')
  sheetTemplate.clear()
  var date = new Date()
  var summaryByJenisKotak = {}
  var summaryByAreaByDocType = {}
  var summaryByDocTypeByAreaByWh = {}

  //ambil sheet
  sheetArea = spreadsheet.getSheetByName('Master Area')
  sheetJenisDokumen = spreadsheet.getSheetByName('Master Jenis Dokumen')
  sheetJenisKotak = spreadsheet.getSheetByName('Master Jenis Kotak')
  sheetGudang = spreadsheet.getSheetByName('Master Gudang')
  sheetJumlahKotak = spreadsheet.getSheetByName('Jumlah Kotak')
  sheetKotak = spreadsheet.getSheetByName('Kotak')
  sheetKotakBaruHilang = spreadsheet.getSheetByName('Riwayat Kotak Baru Hilang')

  //ambil last row
  lrArea = sheetArea.getLastRow()
  lrJenisDokumen = sheetJenisDokumen.getLastRow()
  lrJenisKotak = sheetJenisKotak.getLastRow()
  lrGudang = sheetGudang.getLastRow()
  lrJumlahKotak = sheetJumlahKotak.getLastRow()
  lrKotak = sheetKotak.getLastRow()
  lrKotakBaruHilang = sheetKotakBaruHilang.getLastRow()

  //generate header
  sheetTemplate.getRange(1,1).setValue('Laporan Jumlah Kotak').setFontWeight("bold")
  sheetTemplate.getRange(2,1).setValue('Tanggal Cetak')
  sheetTemplate.getRange(2,2).setValue(Utilities.formatDate(date,"Asia/Bangkok","dd MMMM yyyy HH:mm:ss"))

  //ambil data master
  dataArea = sheetArea.getRange(1,2,lrArea,1).getValues().map(data=>data[0]).filter(data => data !== "")
  dataJenisDokumen = sheetJenisDokumen.getRange(1,1,lrJenisDokumen,1).getValues().map(data=>data[0]).filter(data => data !== "")
  dataJenisKotak = sheetJenisKotak.getRange(1,1,lrJenisKotak,1).getValues().map(data=>data[0]).filter(data => data !== "")
  dataGudang = sheetGudang.getRange(1,1,lrGudang,1).getValues().map(data=>data[0]).filter(data => data !== "")
  dataJumlahKotak = sheetGudang.getRange(2,1,lrJumlahKotak,6).getValues()
  dataKotak = sheetKotak.getRange(2,1,lrKotak-1,9).getValues()
  dataKotakBaruHilang = sheetKotakBaruHilang.getRange(2,2,lrKotakBaruHilang-1,3).getValues()

  cr = 4

  //Proses membuat laporan jumlah per jenis kotak
  dataJenisKotak.forEach(jenisKotak => {
    const dataRow = dataJenisKotak.findIndex(row => row === jenisKotak) + 2
    // console.log(dataRow)
    //mengambil data summary
    summaryByJenisKotak[`${jenisKotak}`] = {}
    summaryByJenisKotak[`${jenisKotak}`]['tersedia'] = sheetJumlahKotak.getRange(dataRow,6).getValues()[0]
    summaryByJenisKotak[`${jenisKotak}`]['hilang'] = sheetJumlahKotak.getRange(dataRow,5).getValues()[0]
    summaryByJenisKotak[`${jenisKotak}`]['rusak'] = sheetJumlahKotak.getRange(dataRow,4).getValues()[0]
    summaryByJenisKotak[`${jenisKotak}`]['terpakai'] = sheetJumlahKotak.getRange(dataRow,3).getValues()[0]
    summaryByJenisKotak[`${jenisKotak}`]['kapasitas'] = sheetJumlahKotak.getRange(dataRow,2).getValues()[0]
    summaryByJenisKotak[`${jenisKotak}`]['per_gudang'] = {}
    dataGudang.forEach(gudang => {
      summaryByJenisKotak[`${jenisKotak}`]['per_gudang'][gudang] = new Object()
      summaryByJenisKotak[`${jenisKotak}`]['per_gudang'][gudang]['terpakai'] = 0
      summaryByJenisKotak[`${jenisKotak}`]['per_gudang'][gudang]['rusak'] = 0
      summaryByJenisKotak[`${jenisKotak}`]['per_gudang'][gudang]['hilang'] = 0
      dataKotak.forEach(kotak => {
        // console.log(kotak[1])
        if (kotak[1]===jenisKotak && kotak[7]===gudang) {
          // console.log(kotak[8])
          switch (kotak[8]) {
            case "Terpakai":
              summaryByJenisKotak[`${jenisKotak}`]['per_gudang'][gudang]['terpakai']++
              break;
            case "Rusak":
              summaryByJenisKotak[`${jenisKotak}`]['per_gudang'][gudang]['rusak']++
              break;
            case "Hilang":
              summaryByJenisKotak[`${jenisKotak}`]['per_gudang'][gudang]['hilang']++
              break;
          }
        }
      })

      dataKotakBaruHilang.forEach(data => {
        if(data[0]===jenisKotak && data[2]===gudang){
          summaryByJenisKotak[`${jenisKotak}`]['per_gudang'][gudang]['hilang']+=parseInt(data[1])
        }
      })
    })
  });
  // console.log(JSON.stringify(summaryByJenisKotak))

  sheetTemplate.getRange(cr, 1).setValue('Jumlah Per Jenis Kardus').setFontWeight("bold")
  cr+=2

  for (const jenisKotak in summaryByJenisKotak) {
    if (!Object.hasOwn(summaryByJenisKotak, jenisKotak)) continue
    
    const element = summaryByJenisKotak[jenisKotak]
    const label = jenisKotak
    
    sheetTemplate.getRange(cr,1).setValue(label).setFontWeight("bold")
    sheetTemplate.getRange(cr+1,1).setValue('Total kapasitas')
    sheetTemplate.getRange(cr+2,1).setValue('Total kotak terpakai')
    sheetTemplate.getRange(cr+3,1).setValue('Total kotak tersedia')
    sheetTemplate.getRange(cr+1,2).setValue(summaryByJenisKotak[jenisKotak]['kapasitas'])
    sheetTemplate.getRange(cr+2,2).setValue(summaryByJenisKotak[jenisKotak]['kapasitas']-summaryByJenisKotak[jenisKotak]['tersedia'])
    sheetTemplate.getRange(cr+3,2).setValue(summaryByJenisKotak[jenisKotak]['tersedia'])
    cr+=4
    //membuat header per gudang
    sheetTemplate.getRange(cr,1).setValue('Gudang').setFontWeight("bold")
    sheetTemplate.getRange(cr,2).setValue('Terpakai').setFontWeight("bold")
    sheetTemplate.getRange(cr,3).setValue('Rusak').setFontWeight("bold")
    sheetTemplate.getRange(cr,4).setValue('Hilang').setFontWeight("bold")
    cr++
    // menuliskan detail jumlah per gudang
    let total = {}
    total['terpakai'] = 0
    total['rusak'] = 0
    total['hilang'] = 0
    for (const gudang in summaryByJenisKotak[jenisKotak]['per_gudang']) {
      if (!Object.hasOwn(summaryByJenisKotak[jenisKotak]['per_gudang'], gudang)) continue
      
      const element = summaryByJenisKotak[jenisKotak]['per_gudang'][gudang]
      const label = gudang

      sheetTemplate.getRange(cr,1).setValue(label)
      sheetTemplate.getRange(cr,2).setValue(element['terpakai'])
      sheetTemplate.getRange(cr,3).setValue(element['rusak'])
      sheetTemplate.getRange(cr,4).setValue(element['hilang'])
      total['terpakai'] += element['terpakai']
      total['rusak'] += element['rusak']
      total['hilang'] += element['hilang']
      cr++
    }
    //menuliskan total detail jumlah
    sheetTemplate.getRange(cr,1).setValue("Total").setFontWeight("bold")
    sheetTemplate.getRange(cr,2).setValue(total['terpakai'])
    sheetTemplate.getRange(cr,3).setValue(total['rusak'])
    sheetTemplate.getRange(cr,4).setValue(total['hilang'])
    cr+=2
  }

  //Proses pembuatan laporan per area per jenis dokumen
  cr = 4
  cc = 6
  sheetTemplate.getRange(cr,cc).setValue("Jumlah Per Area Per Jenis Dokumen").setFontWeight("bold")

  cr += 2
  //Membuat headaer
  sheetTemplate.getRange(cr,cc,2,1).setValue("Area").setFontWeight("bold").merge()
  sheetTemplate.getRange(cr,cc+1,2,1).setValue("Retail").setFontWeight("bold").merge()
  sheetTemplate.getRange(cr,cc+2,2,1).setValue("Non Retail").setFontWeight("bold").merge()
  sheetTemplate.getRange(cr,cc+3,2,1).setValue("Consent Letter").setFontWeight("bold").merge()
  sheetTemplate.getRange(cr,cc+4,2,1).setValue("Total").setFontWeight("bold").merge()
  cr+=2

  //membuatkan default value laporan per area per jenis dokumen
  dataArea.forEach(area => {
    summaryByAreaByDocType[`${area}`] = new Object()
    dataJenisDokumen.forEach(jenisDokumen => {
      summaryByAreaByDocType[`${area}`][`${jenisDokumen}`] = 0
    })
    summaryByAreaByDocType[`${area}`][`total`] = 0
  })

  dataArea.forEach(area=>{
    dataJenisDokumen.forEach(jenisDokumen=>{
      dataKotak.forEach(kotak => {
        if(kotak[2]===jenisDokumen && kotak[5]===area){
          summaryByAreaByDocType[area][jenisDokumen]++
          summaryByAreaByDocType[area]['total']++
        }
      })
    })
  })

  // console.log(summaryByAreaByDocType)
  //menuliskan laporan by area by doc type
  let total = {}
  dataJenisDokumen.forEach(jenisDokumen => {
    total[`${jenisDokumen}`] = 0
  });
  total['total'] = 0
  for (const area in summaryByAreaByDocType) {
    if (!Object.hasOwn(summaryByAreaByDocType, area)) continue
    
    const element = summaryByAreaByDocType[area]
    
    sheetTemplate.getRange(cr,cc).setValue(area)
    var addPoint = 1
    dataJenisDokumen.forEach(jenisDokumen=>{
      sheetTemplate.getRange(cr,cc+addPoint).setValue(summaryByAreaByDocType[area][`${jenisDokumen}`])
      total[`${jenisDokumen}`]+=summaryByAreaByDocType[area][`${jenisDokumen}`]
      addPoint++
    })
    sheetTemplate.getRange(cr,cc+addPoint).setValue(summaryByAreaByDocType[area]['total'])
    total['total']+=summaryByAreaByDocType[area]['total']
    cr++
  }
  // console.log(total)
  sheetTemplate.getRange(cr,cc).setValue('Total').setFontWeight('bold')
  var addPoint = 1
  dataJenisDokumen.forEach(jenisDokumen=>{
    sheetTemplate.getRange(cr,cc+addPoint).setValue(total[`${jenisDokumen}`])
    addPoint++
  })
  sheetTemplate.getRange(cr,cc+addPoint).setValue(total['total'])

  // Proses membuat laporan per jenis dokumen per area per gudang
  cr = 4
  cc = 12
  sheetTemplate.getRange(cr,cc).setValue("Laporan Pemakaian Kardus Per Jenis Dokumen Per Area Per Jenis Dokumen").setFontWeight("bold")
  cr+=2

  var addPointCol = 0
  // Membuat Header
  dataJenisDokumen.forEach(jenisDokumen => {
    sheetTemplate.getRange(cr,cc+addPointCol).setValue(jenisDokumen).setFontWeight('bold')
    sheetTemplate.getRange(cr+1,cc+addPointCol).setValue('Area').setFontWeight('bold')
    addPointCol++
    dataGudang.forEach(gudang => {
      sheetTemplate.getRange(cr+1,cc+addPointCol).setValue(gudang).setFontWeight('bold')
      addPointCol++
    })
    sheetTemplate.getRange(cr+1,cc+addPointCol).setValue("Total").setFontWeight('bold')
    addPointCol+=2
  })

  // Membuat model data default untuk laporan per jenis dokumen per area per gudang
  dataJenisDokumen.forEach(jenisDokumen => {
    summaryByDocTypeByAreaByWh[`${jenisDokumen}`] = new Object()
    dataArea.forEach(area => {
      summaryByDocTypeByAreaByWh[`${jenisDokumen}`][`${area}`] = new Object()
      var total = 0
      dataGudang.forEach(gudang => {
        summaryByDocTypeByAreaByWh[`${jenisDokumen}`][`${area}`][`${gudang}`] = 0  
      })
      summaryByDocTypeByAreaByWh[`${jenisDokumen}`][`${area}`]['total'] = total
    })
  })
  // console.log(summaryByDocTypeByAreaByWh)
  
  // Mengambil data untuk laporan per jenis dokumen per area per gudang
  dataJenisDokumen.forEach(jenisDokumen => {
    dataArea.forEach(area => {
      dataGudang.forEach(gudang=> {
        dataKotak.forEach(kotak => {
          if(kotak[2]===jenisDokumen &&
              kotak[5]===area && 
              kotak[7]===gudang && 
              (
                kotak[8]==='Terpakai' 
                // || kotak[8]==='Hilang' 
                // || kotak[8]==='Rusak' 
              )
            ) {
              summaryByDocTypeByAreaByWh[`${jenisDokumen}`][`${area}`][`${gudang}`] ++
              summaryByDocTypeByAreaByWh[`${jenisDokumen}`][`${area}`][`total`] ++
            }
        })
      })
    })
  })

  // console.log('setelah cek data kotak')
  // console.log(summaryByDocTypeByAreaByWh)

  // Mencetak laporan per jenis dokumen per area per gudang

  dataJenisDokumen.forEach(jenisDokumen => {
    cr = 8
    var sumTotal = new Object()
    dataGudang.forEach(gudang => {
      sumTotal[`${gudang}`] = 0
    })
    sumTotal['total'] = 0
    dataArea.forEach(area => {
      addPointCol = 0
      sheetTemplate.getRange(cr,cc+addPointCol).setValue(area)
      addPointCol++
      dataGudang.forEach(gudang => {
        sheetTemplate.getRange(cr,cc+addPointCol).setValue(summaryByDocTypeByAreaByWh[jenisDokumen][area][gudang])
        sumTotal[`${gudang}`]+=summaryByDocTypeByAreaByWh[jenisDokumen][area][gudang]
        addPointCol++
      })
      sheetTemplate.getRange(cr,cc+addPointCol).setValue(summaryByDocTypeByAreaByWh[jenisDokumen][area]['total'])
      sumTotal['total']+=summaryByDocTypeByAreaByWh[jenisDokumen][area]['total']
      cr++
    })
    // menulis total
    // console.log(cc,addPointCol)
    sheetTemplate.getRange(cr, cc).setValue('Total').setFontWeight('bold')
    addPointCol = 1
    dataGudang.forEach(gudang => {
      // console.log(cc,addPointCol)
      // console.log(gudang)
      sheetTemplate.getRange(cr,cc+addPointCol).setValue(sumTotal[`${gudang}`])
      addPointCol++
    })
    sheetTemplate.getRange(cr,cc+addPointCol).setValue(sumTotal[`total`])
    
    // console.log(sumTotal)
    dataGudang.forEach(gudang => {
      sumTotal[`${gudang}`] = 0
    })
    sumTotal['total'] = 0

    cc+=addPointCol+2
  })

  generateAndSendLink(`Laporan jumlah kotak`, 'Template Box')

  console.log(e)
  if (e && e.triggerUid) {
    const triggerId = e.triggerUid
    deleteTriggerByUid(triggerId)
  }
  // return 'File selesai diproses dan akan dikirim melalui email. Silahkan cek email. Sesungguhnya Tuhan membersamai orang-orang yang sabar :)'
}