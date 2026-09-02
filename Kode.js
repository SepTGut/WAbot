/**
 * Google Apps Script — Webhook Receiver untuk Bot WhatsApp Laporan Progress
 *
 * Fitur Unggulan:
 * 1. LockService — Cegah bentrok/race condition saat banyak laporan masuk bersamaan
 * 2. Reliable Drive Image CDN — Menggunakan lh3.googleusercontent.com agar gambar selalu tampil di sheet
 * 3. Auto-Formatting — Freeze header, auto column width, auto text wrap, & status styling
 * 4. Formula Injection & Number Protection — Mencegah rumus berbahaya & format nomor WA tetap rapi
 * 5. GET endpoint — Health-check saat URL dibuka di browser
 */

// Konfigurasi Token: Boleh hardcoded di sini atau diset via Script Properties
const TOKEN_RAHASIA = PropertiesService.getScriptProperties().getProperty("TOKEN_RAHASIA") || "tewelmambu123";
const NAMA_FOLDER_FOTO = "Foto Laporan Bot";

const KOLOM_HEADER = [
  "Waktu Input", "Nomor Pengirim", "Nama Pengirim", "Gedung", "Tanggal Pengerjaan", "Jam Selesai",
  "Sub Pekerjaan", "Titik Lokasi", "Progres", "Status", "Kendala", "Foto", "Link Foto Drive",
];

/**
 * Health-check endpoint saat URL Web App dibuka di browser
 */
function doGet(e) {
  return jsonResponse({
    status: "online",
    service: "WhatsApp Reporting Bot Webhook",
    timestamp: new Date().toISOString(),
    message: "Webhook siap menerima data POST dari bot."
  });
}

/**
 * Handler utama untuk menerima data laporan dari Bot WhatsApp
 */
function doPost(e) {
  // Gunakan ScriptLock agar write ke spreadsheet antre rapi (mencegah baris tertimpa)
  const lock = LockService.getScriptLock();
  const lockDiterima = lock.tryLock(30000); // Tunggu maks 30 detik

  if (!lockDiterima) {
    return jsonResponse({ ok: false, error: "Server sibuk memproses laporan lain, coba lagi dalam beberapa detik." });
  }

  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ ok: false, error: "Payload kosong (tidak ada data post)." });
    }

    const data = JSON.parse(e.postData.contents);

    // Validasi token keamanan
    if (data.token !== TOKEN_RAHASIA) {
      return jsonResponse({ ok: false, error: "Token rahasia tidak valid." });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Laporan");
    const isNewSheet = !sheet;

    if (isNewSheet) {
      sheet = ss.insertSheet("Laporan");
    }

    // Buat header & format jika sheet baru atau masih kosong
    if (sheet.getLastRow() === 0) {
      inisialisasiSheet(sheet);
    }

    // Sanitasi data agar aman dari formula injection & format nomor tetap terjaga
    const nomorPengirimAman = data.pengirim ? `'${data.pengirim}` : "";
    const namaPengirimAman = sanitasiNilai(data.namaPengirim || "Tidak diketahui");
    const gedungAman = sanitasiNilai(data.gedung || "");
    const tanggalMulaiAman = sanitasiNilai(data.tanggalMulai || "");
    const jamSelesaiAman = sanitasiNilai(data.jamSelesai || "");
    const subPekerjaanAman = sanitasiNilai(data.subPekerjaan || "");
    const titikLokasiAman = sanitasiNilai(data.titikLokasi || "");
    const progresAman = sanitasiNilai(data.progres || "");
    const statusAman = sanitasiNilai(data.status || "Open");
    const kendalaAman = sanitasiNilai(data.kendala || "");
    const waktuInputAman = sanitasiNilai(data.waktuInput || new Date().toLocaleString("id-ID", { timeZone: "Asia/Jakarta" }));

    // Tulis baris data
    sheet.appendRow([
      waktuInputAman,
      nomorPengirimAman,
      namaPengirimAman,
      gedungAman,
      tanggalMulaiAman,
      jamSelesaiAman,
      subPekerjaanAman,
      titikLokasiAman,
      progresAman,
      statusAman,
      kendalaAman,
      "", // Placeholder Foto (akan diisi formula IMAGE)
      "", // Placeholder Link Foto Drive
    ]);

    const barisBaru = sheet.getLastRow();
    const kolomStatus = 10;
    const kolomFoto = 12;
    const kolomLinkFoto = 13;

    // Format visual per baris
    sheet.getRange(barisBaru, 1, 1, KOLOM_HEADER.length).setVerticalAlignment("middle");
    sheet.getRange(barisBaru, 1).setHorizontalAlignment("center"); // Waktu Input
    sheet.getRange(barisBaru, 5, 1, 2).setHorizontalAlignment("center"); // Tanggal & Jam
    sheet.getRange(barisBaru, kolomStatus).setHorizontalAlignment("center"); // Status

    // Warna status (Hijau untuk Close, Kuning/Oranye untuk Open)
    const rangeStatus = sheet.getRange(barisBaru, kolomStatus);
    if (statusAman.toLowerCase() === "close") {
      rangeStatus.setBackground("#D4EDDA").setFontColor("#155724").setFontWeight("bold");
    } else {
      rangeStatus.setBackground("#FFF3CD").setFontColor("#856404").setFontWeight("bold");
    }

    // Upload & Sematkan Foto jika dilampirkan
    if (data.fotoBase64) {
      try {
        const infoFoto = simpanFotoKeGoogleDrive(
          data.fotoBase64,
          data.fotoMime || "image/jpeg",
          data.namaFileFoto || (`foto-${Date.now()}.jpg`)
        );

        // Gunakan CDN URL Google Photos / Drive agar formula IMAGE() stabil
        sheet.getRange(barisBaru, kolomFoto).setFormula(`=IMAGE("${infoFoto.cdnUrl}"; 1)`);
        sheet.getRange(barisBaru, kolomLinkFoto).setFormula(`=HYPERLINK("${infoFoto.viewUrl}"; "Buka Foto Asli ↗")`);
        sheet.setRowHeight(barisBaru, 90);
      } catch (errFoto) {
        console.error("Gagal upload foto:", errFoto);
        sheet.getRange(barisBaru, kolomFoto).setValue("(Gagal upload: " + errFoto.message + ")");
      }
    } else {
      sheet.getRange(barisBaru, kolomFoto).setValue("-");
      sheet.getRange(barisBaru, kolomLinkFoto).setValue("-");
      sheet.setRowHeight(barisBaru, 35);
    }

    return jsonResponse({ ok: true, baris: barisBaru });
  } catch (err) {
    console.error("Error doPost:", err);
    return jsonResponse({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Format & desain awal sheet saat pertama kali dibuat
 */
function inisialisasiSheet(sheet) {
  sheet.appendRow(KOLOM_HEADER);
  const headerRange = sheet.getRange(1, 1, 1, KOLOM_HEADER.length);

  // Styling header: Biru gelap modern dengan teks putih tebal
  headerRange
    .setFontWeight("bold")
    .setBackground("#1A73E8")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  sheet.setRowHeight(1, 40);
  sheet.setFrozenRows(1); // Freeze baris 1 agar selalu terlihat saat scroll

  // Atur lebar kolom ideal
  const lebarKolom = {
    1: 160, // Waktu Input
    2: 150, // Nomor Pengirim
    3: 150, // Nama Pengirim
    4: 160, // Gedung
    5: 120, // Tanggal
    6: 100, // Jam Selesai
    7: 120, // Sub Pekerjaan
    8: 180, // Titik Lokasi
    9: 250, // Progres
    10: 100, // Status
    11: 200, // Kendala
    12: 110, // Foto
    13: 140, // Link Foto Drive
  };

  for (const [col, width] of Object.entries(lebarKolom)) {
    sheet.setColumnWidth(Number(col), width);
  }

  // Wrap text pada kolom progres & kendala
  sheet.getRange("I:I").setWrap(true);
  sheet.getRange("K:K").setWrap(true);
}

/**
 * Simpan foto (base64) ke folder Google Drive dengan CDN URL
 */
function simpanFotoKeGoogleDrive(base64Data, mimeType, namaFile) {
  if (!base64Data || typeof base64Data !== "string" || base64Data.length === 0) {
    throw new Error("Data foto base64 kosong.");
  }

  const folders = DriveApp.getFoldersByName(NAMA_FOLDER_FOTO);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(NAMA_FOLDER_FOTO);

  const mimeTypeAman = String(mimeType || "image/jpeg");
  const namaFileAman = String(namaFile || (`foto-${Date.now()}.jpg`));

  const bytes = Utilities.base64Decode(base64Data);
  if (bytes.length === 0) {
    throw new Error("Hasil decode base64 kosong (0 byte)");
  }

  const blob = Utilities.newBlob(bytes, mimeTypeAman, namaFileAman);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const fileId = file.getId();
  return {
    fileId: fileId,
    viewUrl: file.getUrl(),
    // URL CDN lh3.googleusercontent.com yang paling stabil untuk fungsi IMAGE() di Google Sheets
    cdnUrl: "https://lh3.googleusercontent.com/d/" + fileId
  };
}

/**
 * Sanitasi string mencegah formula injection
 */
function sanitasiNilai(nilai) {
  if (nilai === null || nilai === undefined) return "";
  const str = String(nilai).trim();
  if (/^[=+\-@\t\r]/.test(str)) {
    return "'" + str;
  }
  return str;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Fungsi pembantu untuk trigger izin Drive pertama kali dari editor
 */
function testPermissions() {
  DriveApp.getRootFolder();
  SpreadsheetApp.getActiveSpreadsheet();
  console.log("Izin Drive & Sheets berhasil diverifikasi.");
}
