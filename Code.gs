// ================= GOOGLE APPS SCRIPT BACKEND - AdaNih Presensi =================
// Hubungkan script ini ke Google Sheet Anda dan deploy sebagai Web App

const SPREADSHEET_ID = "YOUR_SPREADSHEET_ID_HERE"; // Ganti dengan ID Sheet Anda
const SHEET_USERS = "users"; // Nama sheet untuk data user
const SHEET_ATTENDANCE = "attendance";
const SHEET_PERMISSIONS = "permissions";

// ====== FUNGSI UTAMA: LOGIN USER DARI GOOGLE SHEETS ======
function loginUser(username, password) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_USERS);
    const data = sheet.getDataRange().getValues();
    
    // Header: [NIP, Nama, Password, Role, Jabatan, Pangkat, Alamat, Foto URL, Atasan, Skema Kerja]
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const nip = String(row[0]).trim();
      const nama = String(row[1]).trim();
      const storedPassword = String(row[2]).trim();
      const role = String(row[3]).toLowerCase().trim();
      const jabatan = String(row[4]).trim();
      const pangkat = String(row[5]).trim();
      const alamat = String(row[6]).trim();
      const foto = String(row[7]).trim();
      const atasan = String(row[8]).trim();
      const skemaKerja = String(row[9]).trim();
      
      // Cek username (NIP) dan password
      if (nip === username && storedPassword === password && nip !== "") {
        return {
          success: true,
          user: {
            username: nip,
            nip: nip,
            nama: nama,
            role: role,
            jabatan: jabatan,
            pangkat: pangkat,
            alamat: alamat,
            foto: foto,
            atasan: atasan,
            skemaKerja: skemaKerja
          }
        };
      }
    }
    
    // Jika tidak ditemukan
    return { 
      success: false, 
      message: "Username atau Password salah!" 
    };
  } catch (error) {
    Logger.log("Error loginUser: " + error);
    return { 
      success: false, 
      message: "Terjadi kesalahan: " + error.toString() 
    };
  }
}

// ====== AMBIL STATUS ABSEN HARI INI ======
function getTodayAttendanceStatus(nip) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_ATTENDANCE);
    const data = sheet.getDataRange().getValues();
    const today = new Date().toISOString().split('T')[0];
    
    // Header: [Tanggal, NIP, Nama, Waktu Masuk, Foto Masuk, Lokasi Masuk, Jarak Masuk, Waktu Pulang, Foto Pulang, Lokasi Pulang, Jarak Pulang, Status]
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (String(row[1]).trim() === nip && String(row[0]).trim() === today) {
        return {
          tanggal: row[0],
          waktu_masuk: row[3],
          waktu_pulang: row[7],
          status: row[11]
        };
      }
    }
    
    return null;
  } catch (error) {
    Logger.log("Error getTodayAttendanceStatus: " + error);
    return null;
  }
}

// ====== AMBIL LIVE PRESENSI SEMUA GURU ======
function getLivePresensi() {
  try {
    const userSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_USERS);
    const attSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_ATTENDANCE);
    
    const userData = userSheet.getDataRange().getValues();
    const attData = attSheet.getDataRange().getValues();
    const today = new Date().toISOString().split('T')[0];
    
    let liveList = [];
    
    for (let i = 1; i < userData.length; i++) {
      const userRow = userData[i];
      const nip = String(userRow[0]).trim();
      const nama = String(userRow[1]).trim();
      const jabatan = String(userRow[4]).trim();
      const foto = String(userRow[7]).trim();
      
      if (nip === "") continue;
      
      // Cari record presensi hari ini
      let attRecord = null;
      for (let j = 1; j < attData.length; j++) {
        if (String(attData[j][1]).trim() === nip && String(attData[j][0]).trim() === today) {
          attRecord = attData[j];
          break;
        }
      }
      
      liveList.push({
        nip: nip,
        nama: nama,
        jabatan: jabatan,
        foto: foto,
        waktu_masuk: attRecord ? attRecord[3] : "-",
        waktu_pulang: attRecord ? attRecord[7] : "-",
        lokasi_masuk: attRecord ? attRecord[5] : "-",
        status: attRecord ? attRecord[11] : "Alfa"
      });
    }
    
    return liveList;
  } catch (error) {
    Logger.log("Error getLivePresensi: " + error);
    return [];
  }
}

// ====== KIRIM PRESENSI (MASUK/PULANG) ======
function submitPresensi(nip, nama, type, lat, lng, photoBase64) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_ATTENDANCE);
    const today = new Date().toISOString().split('T')[0];
    const now = new Date();
    const timeStr = now.toTimeString().split(' ')[0]; // HH:MM:SS
    
    const data = sheet.getDataRange().getValues();
    let found = false;
    
    // Cari record hari ini untuk user ini
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === nip && String(data[i][0]).trim() === today) {
        // Record sudah ada
        if (type === "masuk") {
          return { success: false, message: "Sudah absen masuk!" };
        } else {
          // Update waktu pulang
          sheet.getRange(i + 1, 8).setValue(timeStr); // Waktu Pulang
          sheet.getRange(i + 1, 9).setValue("Foto: " + photoBase64.substring(0, 30) + "...");
          sheet.getRange(i + 1, 10).setValue(lat + "," + lng);
          sheet.getRange(i + 1, 12).setValue("Pulang");
          found = true;
          break;
        }
      }
    }
    
    if (!found && type === "masuk") {
      // Buat record baru
      const newRow = [
        today,
        nip,
        nama,
        timeStr, // Waktu Masuk
        "Foto: " + photoBase64.substring(0, 30) + "...",
        lat + "," + lng,
        "0m",
        "",
        "",
        "",
        "",
        "Hadir"
      ];
      sheet.appendRow(newRow);
      found = true;
    }
    
    if (found) {
      return { success: true, message: "Berhasil mengirim presensi " + type.toUpperCase() + "!" };
    } else {
      return { success: false, message: "Belum absen masuk!" };
    }
  } catch (error) {
    Logger.log("Error submitPresensi: " + error);
    return { success: false, message: "Terjadi kesalahan: " + error.toString() };
  }
}

// ====== KIRIM IZIN ======
function submitIzin(payload) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_PERMISSIONS);
    const idIzin = "IZIN-" + Math.floor(100000 + Math.random() * 900000);
    const now = new Date();
    
    const newRow = [
      idIzin,
      payload.nip,
      payload.nama || "Guru SMAN 1 Parung",
      payload.tipe,
      payload.lokasi_peta || "",
      payload.tanggal_mulai || "",
      payload.tanggal_selesai || "",
      payload.jam_tiba_pulang || "",
      payload.keterangan || "",
      "Menunggu",
      now.toLocaleString()
    ];
    
    sheet.appendRow(newRow);
    return { success: true, message: "Pengajuan izin berhasil terkirim!" };
  } catch (error) {
    Logger.log("Error submitIzin: " + error);
    return { success: false, message: "Terjadi kesalahan: " + error.toString() };
  }
}

// ====== AMBIL DAFTAR IZIN ======
function getIzinList(nip) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_PERMISSIONS);
    const data = sheet.getDataRange().getValues();
    
    let izinList = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (String(row[1]).trim() === nip) {
        izinList.push({
          id_izin: row[0],
          nip: row[1],
          nama: row[2],
          tipe: row[3],
          lokasi_peta: row[4],
          tanggal_mulai: row[5],
          tanggal_selesai: row[6],
          jam_tiba_pulang: row[7],
          keterangan: row[8],
          status_approval: row[9],
          tanggal_input: row[10]
        });
      }
    }
    
    return izinList;
  } catch (error) {
    Logger.log("Error getIzinList: " + error);
    return [];
  }
}

// ====== UPDATE STATUS APPROVAL IZIN ======
function updateIzinStatus(idIzin, statusBaru) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_PERMISSIONS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === idIzin) {
        sheet.getRange(i + 1, 10).setValue(statusBaru); // Status Approval
        return { success: true, message: "Berhasil memperbarui status izin ke: " + statusBaru };
      }
    }
    
    return { success: false, message: "Data izin tidak ditemukan." };
  } catch (error) {
    Logger.log("Error updateIzinStatus: " + error);
    return { success: false, message: "Terjadi kesalahan: " + error.toString() };
  }
}

// ====== AMBIL HISTORIS KEHADIRAN ======
function getHistorisKehadiran(nip) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_ATTENDANCE);
    const data = sheet.getDataRange().getValues();
    
    let history = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === nip) {
        history.push({
          tanggal: data[i][0],
          waktu_masuk: data[i][3],
          waktu_pulang: data[i][7],
          status: data[i][11]
        });
      }
    }
    
    return history;
  } catch (error) {
    Logger.log("Error getHistorisKehadiran: " + error);
    return [];
  }
}

// ====== UPDATE ALAMAT PROFIL ======
function updateProfilAlamat(nip, alamat) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_USERS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === nip) {
        sheet.getRange(i + 1, 7).setValue(alamat); // Alamat
        return { success: true, message: "Alamat profil berhasil diperbarui!" };
      }
    }
    
    return { success: false, message: "User tidak ditemukan." };
  } catch (error) {
    Logger.log("Error updateProfilAlamat: " + error);
    return { success: false, message: "Terjadi kesalahan: " + error.toString() };
  }
}

// ====== AMBIL SEMUA DAFTAR GURU ======
function getAllGuruList() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_USERS);
    const data = sheet.getDataRange().getValues();
    
    let guruList = [];
    for (let i = 1; i < data.length; i++) {
      const nip = String(data[i][0]).trim();
      const nama = String(data[i][1]).trim();
      const role = String(data[i][3]).toLowerCase().trim();
      
      if (nip !== "" && role === "guru") {
        guruList.push({
          nip: nip,
          nama: nama
        });
      }
    }
    
    return guruList;
  } catch (error) {
    Logger.log("Error getAllGuruList: " + error);
    return [];
  }
}

// ====== HELPER: DEPLOY AS WEB APP ======
function doGet(e) {
  return HtmlService.createHtmlOutput("AdaNih API Backend Running");
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    
    let result;
    switch(action) {
      case "loginUser":
        result = loginUser(data.username, data.password);
        break;
      case "getTodayAttendanceStatus":
        result = getTodayAttendanceStatus(data.nip);
        break;
      case "getLivePresensi":
        result = getLivePresensi();
        break;
      case "submitPresensi":
        result = submitPresensi(data.nip, data.nama, data.type, data.lat, data.lng, data.photoBase64);
        break;
      case "submitIzin":
        result = submitIzin(data.payload);
        break;
      case "getIzinList":
        result = getIzinList(data.nip);
        break;
      case "updateIzinStatus":
        result = updateIzinStatus(data.idIzin, data.statusBaru);
        break;
      case "getHistorisKehadiran":
        result = getHistorisKehadiran(data.nip);
        break;
      case "updateProfilAlamat":
        result = updateProfilAlamat(data.nip, data.alamat);
        break;
      case "getAllGuruList":
        result = getAllGuruList();
        break;
      default:
        result = { success: false, message: "Action tidak dikenali" };
    }
    
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    Logger.log("Error doPost: " + error);
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
