// Konfigurasi ID Database dan Folder Drive Anda
const SPREADSHEET_ID = '1aAyfMM_ZyZKeLFyvcyghylDrvkjVGYlQ1Pgwvn9n2pQ';
const FOLDER_ID = '11K0o9xXPvvAxORY5MlhqtERM8W5lhbls';

// JALANKAN FUNGSI INI SEKALI DARI EDITOR UNTUK MENDAPATKAN IZIN EMAIL
function setupPermissions() {
  MailApp.getRemainingDailyQuota();
}

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
    var action = data.action || 'scan'; // Default ke 'scan' jika tidak ada action

    // Buka Spreadsheet untuk mencatat histori scan
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheetLog = ss.getSheetByName("Log_Scan"); 
    if (!sheetLog) { 
      sheetLog = ss.getSheets()[0]; 
    }

    if (action === 'scan') {
      var barcodeText = data.barcodeText;
      var namaMaterial = data.namaMaterial || "";
      var base64Photo = data.base64Photo;
      var photoUrl = "";

      if (base64Photo && base64Photo !== "") {
        var decodedData = Utilities.base64Decode(base64Photo);
        var blob = Utilities.newBlob(decodedData, 'image/jpeg', 'Scan_' + new Date().getTime() + '.jpg');
        var folder = DriveApp.getFolderById(FOLDER_ID);
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        photoUrl = file.getUrl();
      }

      var timestamp = new Date();
      sheetLog.appendRow([timestamp, barcodeText, namaMaterial, photoUrl, "Tercatat Otomatis"]);

      var response = { status: 'success', message: 'Data berhasil disimpan!' };
      return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
      
    } else if (action === 'fetch_today') {
      var rawData = sheetLog.getDataRange().getValues();
      var displayData = sheetLog.getDataRange().getDisplayValues();
      var today = new Date();
      var resultData = [];
      
      // Looping data, mulai dari indeks 1 untuk melewati Header
      for(var i = 1; i < rawData.length; i++) {
        var rowDate = rawData[i][0];
        if (rowDate instanceof Date) {
          if (rowDate.getDate() === today.getDate() && 
              rowDate.getMonth() === today.getMonth() && 
              rowDate.getFullYear() === today.getFullYear()) {
            
            resultData.push({
              timestamp: displayData[i][0],
              barcode: displayData[i][1],
              nama: displayData[i][2],
              status: displayData[i][4]
            });
          }
        }
      }
      
      var response = { status: 'success', data: resultData };
      return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);

    } else if (action === 'send_email') {
      var emailTo = data.emailTo;
      var subject = data.subject;
      var tableData = data.tableData;

      // Buat struktur tabel HTML
      var htmlBody = "<h3 style='font-family: sans-serif;'>Laporan Scan Barcode</h3>";
      htmlBody += "<table border='1' cellpadding='8' style='border-collapse: collapse; font-family: sans-serif; width: 100%;'>";
      htmlBody += "<tr style='background-color: #f1f5f9;'><th>Waktu</th><th>Barcode</th><th>Nama Material</th><th>Status</th></tr>";
      
      for(var j = 0; j < tableData.length; j++) {
        htmlBody += "<tr>";
        htmlBody += "<td>" + tableData[j].timestamp + "</td>";
        htmlBody += "<td>" + tableData[j].barcode + "</td>";
        htmlBody += "<td>" + tableData[j].nama + "</td>";
        htmlBody += "<td>" + tableData[j].status + "</td>";
        htmlBody += "</tr>";
      }
      htmlBody += "</table>";

      MailApp.sendEmail({ to: emailTo, subject: subject, htmlBody: htmlBody });

      var response = { status: 'success', message: 'Email berhasil dikirim!' };
      return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
    }

  } catch (error) {
    var response = { status: 'error', message: error.toString() };
    return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
  }
}
