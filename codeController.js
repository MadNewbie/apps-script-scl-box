function doPost(params) {
  let rawData = {}
  if(params.postData && params.postData.type === 'application/json'){
    rawData = JSON.parse(params.postData.contents)
  } else if (params.parameter) {
    rawData = params.parameter
  }

  const action = rawData.action
  const data = {...rawData}
  switch (action) {
    case 'kbt':
      return kotakBaruTambah(data)
      break;

    case 'kbh':
      return kotakBaruHilang(data)
      break;

    case 'kth':
      return kotakTambah(data)
      break;
    
    case 'kuh':
      return kotakUbah(data)
      break;

    case 'khg':
      return kotakHilang(data)
      break;

    case 'krk':
      return kotakRusak(data)
      break;

    case 'kpj':
      return kotakPinjam(data)
      break;

    case 'kpk':
      return kotakPengembalianKotak(data)
      break;

    case 'ktf':
      return kotakTransfer(data)
      break;
  
    default:
      return jsonResponse({
        'message': "Aksi belum terdaftar, silahkan menghubungi pengembang"
      })
      break;
  }
}

function doGet() {
  return jsonResponse({
    'message': 'Google Script Spreadsheet API aktif',
  })
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON)
}