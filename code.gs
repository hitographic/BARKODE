// Konfigurasi ID Database dan Folder Drive Anda
const SPREADSHEET_ID = '1aAyfMM_ZyZKeLFyvcyghylDrvkjVGYlQ1Pgwvn9n2pQ';
const FOLDER_ID = '11K0o9xXPvvAxORY5MlhqtERM8W5lhbls';

// Fungsi untuk MENGIRIM data master ke Frontend saat halaman dimuat
function doGet(e) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    // Asumsi: Data Master Material ada di sheet pertama (kiri sendiri)
    // Pastikan Kolom A = Nama Material, Kolom B = Kode Barcode
    var sheet = ss.getSheets()[0]; 
    var data = sheet.getDataRange().getDisplayValues(); // Gunakan getDisplayValues agar angka dibaca sebagai teks persis seperti di layar
    
    var masterData = [];
    // Looping dari baris ke-2 (indeks 1) untuk melewati baris judul (header)
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][1]) { // Pastikan baris tidak kosong
        masterData.push({
          nama: String(data[i][0]).trim(), // Bersihkan spasi berlebih
          kode: String(data[i][1]).trim()  // Bersihkan spasi berlebih
        });
      }
    }
    
    var response = { status: 'success', data: masterData };
    return ContentService.createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: error.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Fungsi untuk MENERIMA data scan dan foto dari Frontend
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var barcodeText = data.barcodeText;
    var namaMaterial = data.namaMaterial || "";
    var base64Photo = data.base64Photo;
    var photoUrl = "";

    // Jika ada foto, proses fotonya (sekarang menjadi opsional/dilewati)
    if (base64Photo && base64Photo !== "") {
      var decodedData = Utilities.base64Decode(base64Photo);
      var blob = Utilities.newBlob(decodedData, 'image/jpeg', 'Scan_' + new Date().getTime() + '.jpg');
      var folder = DriveApp.getFolderById(FOLDER_ID);
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      photoUrl = file.getUrl();
    }

    // Buka Spreadsheet untuk mencatat histori scan
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Asumsi: Histori log scan dicatat di sheet kedua. 
    // Jika Anda hanya punya 1 sheet, ubah jadi ss.getSheets()[0]
    // Dianjurkan buat Sheet/Tab baru bernama "Log_Scan" di Google Sheets Anda
    var sheetLog = ss.getSheetByName("Log_Scan"); 
    if (!sheetLog) { 
      // Jika tab Log_Scan belum dibuat, gunakan tab pertama
      sheetLog = ss.getSheets()[0]; 
    }

    var timestamp = new Date();
    // Susunan Kolom Baru: [Waktu, Hasil Scan, Nama Material, Link Foto, Status]
    sheetLog.appendRow([timestamp, barcodeText, namaMaterial, photoUrl, "Tercatat Otomatis"]);

    var response = {
      status: 'success',
      message: 'Data berhasil disimpan!',
      photoUrl: photoUrl
    };
    
    return ContentService.createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    var response = {
      status: 'error',
      message: error.toString()
    };
    return ContentService.createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
