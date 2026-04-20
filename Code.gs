/**
 * KARIMUNKU TERANG - Backend Logic
 * Mengatur penyimpanan data, Login, Role, Histori, Foto, Teknisi, dan Notifikasi Email
 */

const FOLDER_ID = "1bJfoalUY1fqJ-1bgKRGQhdkTAC-C3fWr"; // ID Folder Google Drive Anda
const SHEET_NAME = "Data Laporan";
const USERS_SHEET = "Data Pengguna"; 
const ADMIN_EMAIL = "riedhosatria@gmail.com"; // Email yang akan menerima notifikasi laporan baru

function doGet(e) {
  try {
    var html = HtmlService.createHtmlOutputFromFile('Index');
    html.setTitle('KARIMUNKU TERANG');
    html.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    html.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    return html;
  } catch (err) {
    try {
      var htmlLower = HtmlService.createHtmlOutputFromFile('index');
      htmlLower.setTitle('KARIMUNKU TERANG');
      htmlLower.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      htmlLower.addMetaTag('viewport', 'width=device-width, initial-scale=1');
      return htmlLower;
    } catch (err2) {
      return HtmlService.createHtmlOutput("<div style='padding:40px; font-family:sans-serif;'><h1>⚠️ Gagal Memuat Tampilan</h1><p>Pastikan nama file HTML Anda sudah benar, yaitu <b>Index.html</b> atau <b>index.html</b>.</p></div>");
    }
  }
}

function getOrCreateSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("Script tidak terhubung ke Spreadsheet.");
  
  let sheetLaporan = ss.getSheetByName(SHEET_NAME);
  const headersLaporan = ["ID LAPORAN", "WAKTU", "NAMA PELAPOR", "WHATSAPP", "NIK", "KECAMATAN", "ALAMAT DETAIL", "KATEGORI", "DESKRIPSI", "LATITUDE", "LONGITUDE", "URL MAPS", "URL FOTO", "STATUS", "HISTORI STATUS", "FOTO PEKERJAAN", "FOTO SELESAI", "TINDAKAN MATERIAL", "TEKNISI PELAKSANA"];
  
  if (!sheetLaporan) {
    sheetLaporan = ss.insertSheet(SHEET_NAME);
    sheetLaporan.getRange(1, 1, 1, headersLaporan.length).setValues([headersLaporan]).setBackground("#f39c12").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
    sheetLaporan.setFrozenRows(1);
    sheetLaporan.autoResizeColumns(1, headersLaporan.length);
  } else {
     if (sheetLaporan.getLastColumn() === 0) {
         sheetLaporan.getRange(1, 1, 1, headersLaporan.length).setValues([headersLaporan]).setBackground("#f39c12").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
     } else {
         let currentHeaders = sheetLaporan.getRange(1, 1, 1, sheetLaporan.getLastColumn()).getValues()[0];
         for (let i = 0; i < headersLaporan.length; i++) {
             if (currentHeaders.indexOf(headersLaporan[i]) === -1) {
                 sheetLaporan.getRange(1, currentHeaders.length + 1).setValue(headersLaporan[i]).setBackground("#f39c12").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
                 currentHeaders.push(headersLaporan[i]);
             }
         }
     }
  }

  let sheetUsers = ss.getSheetByName(USERS_SHEET);
  if (!sheetUsers) {
    sheetUsers = ss.insertSheet(USERS_SHEET);
    const headersUsers = ["NAMA LENGKAP", "USERNAME", "PASSWORD", "ROLE"];
    sheetUsers.getRange(1, 1, 1, 4).setValues([headersUsers]).setBackground("#2980b9").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
    sheetUsers.setFrozenRows(1);
  }
  
  const defaultUsers = [
     ["Super Administrator", "superadmin", "PjuKarimun@2026Aman!", "Super Admin"],
     ["Teknisi Lapangan 1", "teknisi", "teknisi123", "Teknisi"],
     ["Teknisi Lapangan 2", "teknisi2", "teknisi123", "Teknisi"],
     ["Teknisi Lapangan 3", "teknisi3", "teknisi123", "Teknisi"],
     ["Teknisi Lapangan 4", "teknisi4", "teknisi123", "Teknisi"]
  ];
  
  if (sheetUsers.getLastRow() === 0) {
     const headersUsers = ["NAMA LENGKAP", "USERNAME", "PASSWORD", "ROLE"];
     sheetUsers.appendRow(headersUsers);
     defaultUsers.forEach(u => sheetUsers.appendRow(u));
  } else {
     const data = sheetUsers.getDataRange().getValues();
     const existingUsernames = data.map(row => String(row[1]).trim().toLowerCase());
     defaultUsers.forEach(user => {
       if (!existingUsernames.includes(user[1].toLowerCase())) {
         sheetUsers.appendRow(user);
       }
     });
  }
}

function loginUser(username, password) {
  try {
    getOrCreateSheets(); 
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USERS_SHEET);
    if (!sheet) return { success: false, message: "Sistem database belum siap." };
    
    const data = sheet.getDataRange().getValues();
    const inputUser = String(username || "").trim().toLowerCase();
    const inputPass = String(password || "").trim();
    
    for (let i = 1; i < data.length; i++) {
      const dbUser = String(data[i][1] || "").trim().toLowerCase();
      const dbPass = String(data[i][2] || "").trim();
      
      if (dbUser === inputUser && dbPass === inputPass) {
        return { success: true, nama: String(data[i][0] || ""), role: String(data[i][3] || "") };
      }
    }
    return { success: false, message: "Username atau Password salah!" };
  } catch(e) {
    return { success: false, message: "Error DB: " + e.toString() };
  }
}

function registerUser(nama, username, password) {
  try {
    getOrCreateSheets();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USERS_SHEET);
    if (!sheet) return { success: false, message: "Sistem database belum siap." };
    
    const data = sheet.getDataRange().getValues();
    const inputUser = String(username || "").trim().toLowerCase();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1] || "").trim().toLowerCase() === inputUser) {
        return { success: false, message: "Username sudah terdaftar!" };
      }
    }
    sheet.appendRow([nama, username, password, "Pelapor"]);
    return { success: true, message: "Pendaftaran berhasil! Silakan Login." };
  } catch(e) {
    return { success: false, message: "Error DB: " + e.toString() };
  }
}

function processForm(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    let fileUrl = "Tidak ada foto";
    if (data.fileBase64) {
      try {
        const folder = DriveApp.getFolderById(FOLDER_ID);
        const blob = Utilities.newBlob(Utilities.base64Decode(data.fileBase64), 'image/jpeg', `PJU_${new Date().getTime()}.jpg`);
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fileUrl = file.getUrl();
      } catch (fError) {
        fileUrl = "Gagal upload: " + fError.toString();
      }
    }

    const reportId = "LPR-" + Math.random().toString(36).substr(2, 9).toUpperCase();
    const currentTime = new Date();
    const mapsLinkFormula = `=HYPERLINK("http://maps.google.com/maps?q=${data.lat},${data.lng}", "📍 Buka Peta")`;

    const tz = Session.getScriptTimeZone();
    const waktuFormat = Utilities.formatDate(currentTime, tz, "yyyy-MM-dd'T'HH:mm:ss");
    const historiAwal = JSON.stringify([{ status: "Menunggu Penanganan", waktu: waktuFormat }]);

    sheet.appendRow([
      reportId, currentTime, data.nama, "'" + data.whatsapp, "'" + (data.nik || "-"),
      data.wilayah, data.alamat, data.kategori_kerusakan, data.detail_kerusakan,
      data.lat, data.lng, mapsLinkFormula, fileUrl, "Menunggu Penanganan", historiAwal, "", "", "[]", ""
    ]);

    // ==========================================
    // KODE BARU: KIRIM NOTIFIKASI EMAIL KE ADMIN
    // ==========================================
    try {
      const emailSubject = `🚨 [PJU Baru] Laporan Masuk: ${reportId} - Kec. ${data.wilayah}`;
      const emailBody = `
        <div style="font-family: Arial, sans-serif; color: #333; max-width: 600px; margin: auto; border: 1px solid #ddd; border-radius: 10px; overflow: hidden;">
          <div style="background-color: #f39c12; color: #fff; padding: 20px; text-align: center;">
            <h2 style="margin: 0;">KARIMUNKU TERANG</h2>
            <p style="margin: 5px 0 0 0;">Laporan Kerusakan PJU Baru</p>
          </div>
          <div style="padding: 20px;">
            <p>Halo Admin,</p>
            <p>Terdapat laporan kerusakan PJU baru yang baru saja dikirimkan oleh warga. Berikut adalah detail laporannya:</p>
            <table style="width: 100%; border-collapse: collapse; margin-top: 15px;">
              <tr><td style="padding: 8px; border-bottom: 1px solid #eee;"><b>ID Laporan</b></td><td style="padding: 8px; border-bottom: 1px solid #eee;">: ${reportId}</td></tr>
              <tr><td style="padding: 8px; border-bottom: 1px solid #eee;"><b>Waktu Laporan</b></td><td style="padding: 8px; border-bottom: 1px solid #eee;">: ${Utilities.formatDate(currentTime, tz, "dd/MM/yyyy HH:mm")}</td></tr>
              <tr><td style="padding: 8px; border-bottom: 1px solid #eee;"><b>Nama Pelapor</b></td><td style="padding: 8px; border-bottom: 1px solid #eee;">: ${data.nama}</td></tr>
              <tr><td style="padding: 8px; border-bottom: 1px solid #eee;"><b>No. WhatsApp</b></td><td style="padding: 8px; border-bottom: 1px solid #eee;">: <a href="https://wa.me/${data.whatsapp.replace(/^0/, '62')}">${data.whatsapp}</a></td></tr>
              <tr><td style="padding: 8px; border-bottom: 1px solid #eee;"><b>Kecamatan</b></td><td style="padding: 8px; border-bottom: 1px solid #eee;">: ${data.wilayah}</td></tr>
              <tr><td style="padding: 8px; border-bottom: 1px solid #eee;"><b>Patokan Alamat</b></td><td style="padding: 8px; border-bottom: 1px solid #eee;">: ${data.alamat}</td></tr>
              <tr><td style="padding: 8px; border-bottom: 1px solid #eee;"><b>Kategori Masalah</b></td><td style="padding: 8px; border-bottom: 1px solid #eee;">: <strong style="color: #c0392b;">${data.kategori_kerusakan}</strong></td></tr>
              <tr><td style="padding: 8px; border-bottom: 1px solid #eee;"><b>Deskripsi</b></td><td style="padding: 8px; border-bottom: 1px solid #eee;">: ${data.detail_kerusakan}</td></tr>
            </table>
            
            <div style="text-align: center; margin-top: 30px;">
              <a href="http://maps.google.com/maps?q=${data.lat},${data.lng}" style="background-color: #2980b9; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; margin-right: 10px; display: inline-block;">📍 Buka Google Maps</a>
              <a href="${fileUrl}" style="background-color: #7f8c8d; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; display: inline-block;">📸 Lihat Foto Laporan</a>
            </div>
            <br>
            <p style="text-align: center; font-size: 0.9em; color: #7f8c8d;">Silakan segera instruksikan Teknisi Lapangan untuk menindaklanjuti laporan ini.</p>
          </div>
          <div style="background-color: #f9f9f9; padding: 15px; text-align: center; font-size: 0.8em; color: #aaa;">
            Email notifikasi ini dibuat otomatis oleh Sistem KARIMUNKU TERANG.
          </div>
        </div>
      `;

      MailApp.sendEmail({
        to: ADMIN_EMAIL,
        subject: emailSubject,
        htmlBody: emailBody
      });
    } catch (emailError) {
      console.error("Gagal mengirim notifikasi email: " + emailError);
    }
    // ==========================================

    return { status: "success", id: reportId };
  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}

function getSafeString(val) {
    if (val === undefined || val === null || val === "Invalid Date") return "";
    if (val instanceof Date) {
        return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
    }
    let strVal = String(val).trim();
    if (strVal === "Invalid Date") return ""; 
    let dateMatch = strVal.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (dateMatch) {
        let timeMatch = strVal.match(/(\d{2}:\d{2})/);
        let timeStr = timeMatch ? timeMatch[1] : "00:00";
        return `${dateMatch[3]}-${dateMatch[2].padStart(2, '0')}-${dateMatch[1].padStart(2, '0')}T${timeStr}:00`;
    }
    return strVal;
}

function getSemuaLaporan() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if(!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    const headers = data[0].map(h => String(h).trim());
    let result = [];
    
    for(let i = 1; i < data.length; i++) {
      result.push({
        id: getSafeString(data[i][headers.indexOf("ID LAPORAN")]),
        waktu: getSafeString(data[i][headers.indexOf("WAKTU")]),
        pelapor: getSafeString(data[i][headers.indexOf("NAMA PELAPOR")]),
        kecamatan: getSafeString(data[i][headers.indexOf("KECAMATAN")]),
        alamat: getSafeString(data[i][headers.indexOf("ALAMAT DETAIL")]),
        kategori: getSafeString(data[i][headers.indexOf("KATEGORI")]),
        status: getSafeString(data[i][headers.indexOf("STATUS")]),
        urlFoto: getSafeString(data[i][headers.indexOf("URL FOTO")]),
        fotoPekerjaan: getSafeString(data[i][headers.indexOf("FOTO PEKERJAAN")]),
        fotoSelesai: getSafeString(data[i][headers.indexOf("FOTO SELESAI")]),
        histori: getSafeString(data[i][headers.indexOf("HISTORI STATUS")]) || "[]",
        tindakan: getSafeString(data[i][headers.indexOf("TINDAKAN MATERIAL")]) || "[]",
        teknisi: getSafeString(data[i][headers.indexOf("TEKNISI PELAKSANA")]),
        lat: getSafeString(data[i][headers.indexOf("LATITUDE")]),
        lng: getSafeString(data[i][headers.indexOf("LONGITUDE")])
      });
    }
    return result.reverse(); 
  } catch (e) {
    return [];
  }
}

function updateLaporanStatus(idLaporan, statusBaru, base64Foto, tindakanArr, namaTeknisi) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) return { success: false, message: "Sheet tidak ditemukan" };

    const data = sheet.getDataRange().getValues();
    let currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    const idIndex = currentHeaders.indexOf("ID LAPORAN");
    const statusIndex = currentHeaders.indexOf("STATUS");
    const historiIndex = currentHeaders.indexOf("HISTORI STATUS");
    let fProsesIdx = currentHeaders.indexOf("FOTO PEKERJAAN");
    let fSelesaiIdx = currentHeaders.indexOf("FOTO SELESAI");
    let tindakanIdx = currentHeaders.indexOf("TINDAKAN MATERIAL");
    let teknisiIdx = currentHeaders.indexOf("TEKNISI PELAKSANA");

    if (fProsesIdx === -1) { fProsesIdx = currentHeaders.length; sheet.getRange(1, fProsesIdx + 1).setValue("FOTO PEKERJAAN").setBackground("#f39c12").setFontColor("#ffffff").setFontWeight("bold"); currentHeaders.push("FOTO PEKERJAAN"); }
    if (fSelesaiIdx === -1) { fSelesaiIdx = currentHeaders.length; sheet.getRange(1, fSelesaiIdx + 1).setValue("FOTO SELESAI").setBackground("#f39c12").setFontColor("#ffffff").setFontWeight("bold"); currentHeaders.push("FOTO SELESAI"); }
    if (tindakanIdx === -1) { tindakanIdx = currentHeaders.length; sheet.getRange(1, tindakanIdx + 1).setValue("TINDAKAN MATERIAL").setBackground("#f39c12").setFontColor("#ffffff").setFontWeight("bold"); currentHeaders.push("TINDAKAN MATERIAL"); }
    if (teknisiIdx === -1) { teknisiIdx = currentHeaders.length; sheet.getRange(1, teknisiIdx + 1).setValue("TEKNISI PELAKSANA").setBackground("#f39c12").setFontColor("#ffffff").setFontWeight("bold"); currentHeaders.push("TEKNISI PELAKSANA"); }

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idIndex]) === String(idLaporan)) {
        
        if (base64Foto) {
          try {
            const folder = DriveApp.getFolderById(FOLDER_ID);
            const namaPrefix = statusBaru === 'Proses Perbaikan' ? 'PROSES' : 'SELESAI';
            const blob = Utilities.newBlob(Utilities.base64Decode(base64Foto), 'image/jpeg', `${namaPrefix}_${idLaporan}.jpg`);
            const file = folder.createFile(blob);
            file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            
            if (statusBaru === 'Proses Perbaikan' && fProsesIdx !== -1) {
              sheet.getRange(i + 1, fProsesIdx + 1).setValue(file.getUrl());
            } else if (statusBaru === 'Selesai' && fSelesaiIdx !== -1) {
              sheet.getRange(i + 1, fSelesaiIdx + 1).setValue(file.getUrl());
            }
          } catch (e) {
            return { success: false, message: "Gagal upload foto Google Drive." };
          }
        }

        if (statusBaru === 'Selesai' && tindakanArr && tindakanIdx !== -1) {
            sheet.getRange(i + 1, tindakanIdx + 1).setValue(JSON.stringify(tindakanArr));
        }

        if (namaTeknisi && teknisiIdx !== -1) {
            sheet.getRange(i + 1, teknisiIdx + 1).setValue(namaTeknisi);
        }

        if (statusIndex !== -1) {
            sheet.getRange(i + 1, statusIndex + 1).setValue(statusBaru);
        }
        
        if (historiIndex !== -1) {
          let historiLama = getSafeString(data[i][historiIndex]);
          let historiArr = [];
          try { historiArr = JSON.parse(historiLama); } catch(e) { historiArr = []; }
          
          const tz = Session.getScriptTimeZone();
          const waktuFormat = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd'T'HH:mm:ss");
          
          historiArr.push({ status: statusBaru, waktu: waktuFormat, teknisi: namaTeknisi || "" });
          sheet.getRange(i + 1, historiIndex + 1).setValue(JSON.stringify(historiArr));
        }

        return { success: true, message: "Status Berhasil Diperbarui" };
      }
    }
    return { success: false, message: "ID Laporan tidak ditemukan" };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function generateAIDescription(lat, lng, wilayah, promptText) {
  const apiKey = "AIzaSyDo_GUcJmZOTCjIOwc_h5JrlFyEPDdWmbU"; 
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key=${apiKey}`;
  const payload = { contents: [{ parts: [{ text: promptText }] }] };
  try {
    const response = UrlFetchApp.fetch(url, { method: "post", contentType: "application/json", payload: JSON.stringify(payload) });
    return JSON.parse(response.getContentText()).candidates[0].content.parts[0].text;
  } catch (e) {
    return "Gagal menyusun deskripsi AI.";
  }
}

function cekStatusLaporan(idLaporan) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) return { ditemukan: false };
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][headers.indexOf("ID LAPORAN")]) === String(idLaporan)) {
        return {
          ditemukan: true,
          id: getSafeString(data[i][headers.indexOf("ID LAPORAN")]),
          waktu: getSafeString(data[i][headers.indexOf("WAKTU")]),
          kecamatan: getSafeString(data[i][headers.indexOf("KECAMATAN")]),
          kategori: getSafeString(data[i][headers.indexOf("KATEGORI")]),
          status: getSafeString(data[i][headers.indexOf("STATUS")]),
          histori: getSafeString(data[i][headers.indexOf("HISTORI STATUS")]) || "[]",
          tindakan: getSafeString(data[i][headers.indexOf("TINDAKAN MATERIAL")]) || "[]",
          teknisi: getSafeString(data[i][headers.indexOf("TEKNISI PELAKSANA")]),
          fotoPekerjaan: getSafeString(data[i][headers.indexOf("FOTO PEKERJAAN")]),
          fotoSelesai: getSafeString(data[i][headers.indexOf("FOTO SELESAI")])
        };
      }
    }
    return { ditemukan: false };
  } catch (e) {
    return { ditemukan: false, error: e.toString() };
  }
}

function getRiwayatPelapor(namaPelapor) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if(!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    let result = [];
    
    for(let i = 1; i < data.length; i++) {
      if (String(data[i][headers.indexOf("NAMA PELAPOR")] || "").trim() === String(namaPelapor).trim()) {
        result.push({
          id: getSafeString(data[i][headers.indexOf("ID LAPORAN")]),
          waktu: getSafeString(data[i][headers.indexOf("WAKTU")]),
          kategori: getSafeString(data[i][headers.indexOf("KATEGORI")]),
          status: getSafeString(data[i][headers.indexOf("STATUS")])
        });
      }
    }
    return result.reverse(); 
  } catch (e) {
    return [];
  }
}