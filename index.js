/**
 * Bot WhatsApp Laporan Progress Pekerjaan
 * Alur: Gedung -> Tanggal Mulai -> Tanggal Selesai -> Sub Pekerjaan (Mekanikal/Elektrikal)
 *       -> Titik Lokasi -> Progres (teks bebas) -> Status (Open/Close) -> Kendala
 *
 * Data Titik Lokasi per Gedung+Sub Pekerjaan diambil dari data-titik-lokasi.xlsx
 * (kolom: Nama Gedung, Sub Pekerjaan, Titik Lokasi, Keterangan)
 *
 * Hasil laporan otomatis tersimpan ke laporan-progress.csv
 *
 * Cara pakai singkat:
 * 1. Siapkan file data-titik-lokasi.xlsx di folder yang sama (lihat template)
 * 2. Jalankan: npm install
 * 3. Jalankan: npm start
 * 4. Scan QR code yang muncul di terminal pakai WhatsApp (menu Perangkat Tertaut)
 */

require("dotenv").config();

const {
  default: makeWASocket,
  useMultiFileAuthState,
  DisconnectReason,
  fetchLatestBaileysVersion,
  downloadMediaMessage,
} = require("@whiskeysockets/baileys");
const qrcode = require("qrcode-terminal");
const pino = require("pino");
const XLSX = require("xlsx");
const fs = require("fs");
const fsPromises = require("fs").promises;
const path = require("path");
const ExcelJS = require("exceljs");

// ====== LOGGING KE FILE ======
// Semua console.log/error otomatis juga tercatat ke bot.log supaya log tidak hilang
// saat terminal ditutup. File di-rotate manual (hapus/rename kalau sudah terlalu besar).
const LOG_FILE_PATH = "bot.log";
const logStream = fs.createWriteStream(LOG_FILE_PATH, { flags: "a" }); // append mode
const _consoleLog = console.log.bind(console);
const _consoleError = console.error.bind(console);
const _consoleWarn = console.warn.bind(console);

function stempelWaktu() {
  return new Date().toLocaleString("id-ID", { timeZone: "Asia/Jakarta" });
}

console.log = (...args) => {
  _consoleLog(...args);
  logStream.write(`[${stempelWaktu()}] [INFO] ${args.map(String).join(" ")}\n`);
};
console.error = (...args) => {
  _consoleError(...args);
  logStream.write(`[${stempelWaktu()}] [ERROR] ${args.map(String).join(" ")}\n`);
};
console.warn = (...args) => {
  _consoleWarn(...args);
  logStream.write(`[${stempelWaktu()}] [WARN] ${args.map(String).join(" ")}\n`);
};

// ====== KONFIGURASI ======
const DATA_FILE_PATH = "data-titik-lokasi.xlsx";
const LAPORAN_FILE_PATH = "laporan-progress.xlsx";
const LAPORAN_CADANGAN_PATH = "laporan-progress-cadangan.xlsx";
const FOLDER_FOTO = "foto-laporan";
const BATAS_PANJANG_INPUT = 500; // karakter maksimal untuk input teks bebas (progres, kendala)
const COOLDOWN_MS = 1000; // jeda minimum antar pesan dari pengirim yang sama (anti-flood)

// Google Sheets via Apps Script Web App (opsional). Kosongkan APPS_SCRIPT_URL
// kalau belum mau setup -- bot otomatis tetap simpan ke Excel lokal saja.
// Nilai diambil dari file .env (lihat .env.example untuk template).
const APPS_SCRIPT_URL = process.env.APPS_SCRIPT_URL || "";
const APPS_SCRIPT_TOKEN = process.env.APPS_SCRIPT_TOKEN || "";

const KOLOM_GEDUNG = "Nama Gedung";
const KOLOM_SUB_PEKERJAAN = "Sub Pekerjaan";
const KOLOM_TITIK_LOKASI = "Titik Lokasi";

// Daftar gedung fallback (urutan preferensi untuk menu). Gedung tambahan yang
// ditemukan di data-titik-lokasi.xlsx tapi belum ada di sini akan ditambahkan
// otomatis di akhir daftar saat bot start.
const DAFTAR_GEDUNG_DASAR = [
  "G.GSG", "Rumah Pompa", "Power House", "Gedung Dapur dan Gudang",
  "Rusun Guru 1 Lt. 1", "Rusun Guru 1 Lt. 2", "Rusun Guru 2 Lt. 1", "Rusun Guru 2 Lt. 2",
  "G.SMP Lt.1", "G.SMP Lt.2",
  "G.Asrama Putri SD1 Lt.1", "G.Asrama Putri SD1 Lt.2",
  "G.Asrama Putra SD1 Lt.1", "G.Asrama Putra SD1 Lt.2",
  "G.Asrama Putri SD2 Lt.1", "G.Asrama Putri SD2 Lt.2",
  "G.Asrama Putra SD2 Lt.1", "G.Asrama Putra SD2 Lt.2",
  "Kantin SD", "Kantin SMP", "Kantin SMA",
  "G.Asrama Putra SMP Lt.1", "G.Asrama Putra SMP Lt.2",
  "G.Asrama Putri SMP Lt.1", "G.Asrama Putri SMP Lt.2",
  "Kawasan", "G. SD Lt.1", "G. SD Lt.2",
  "Gedung Masjid", "Gedung Ibadah",
  "Gedung Guesthouse Lt.1", "Gedung Guesthouse Lt.2",
  "G.Asrama Putra SMA Lt.1", "G.Asrama Putra SMA Lt.2",
  "G.Asrama Putri SMA Lt.1", "G.Asrama Putri SMA Lt.2",
];

// ====== FUNGSI: Baca data Titik Lokasi dari spreadsheet ======
function bacaDataTitikLokasi() {
  if (!fs.existsSync(DATA_FILE_PATH)) {
    console.warn(`⚠️  File ${DATA_FILE_PATH} tidak ditemukan. Bot akan jalan tanpa data titik lokasi.`);
    return [];
  }
  try {
    const workbook = XLSX.readFile(DATA_FILE_PATH);
    const sheetPertama = workbook.SheetNames[0];
    // baris 1 = legenda, baris 2 = header -> data mulai baris ke-3 (index 1)
    // defval: null supaya sel kosong tetap muncul sebagai properti (null), bukan hilang
    const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetPertama], { range: 1, defval: null });

    // Fill-down: baris yang Nama Gedung / Sub Pekerjaan-nya kosong mewarisi nilai dari baris terakhir yang terisi
    let gedungTerakhir = null;
    let subTerakhir = null;
    const dataDiisi = data.map((baris) => {
      if (baris[KOLOM_GEDUNG]) gedungTerakhir = baris[KOLOM_GEDUNG];
      if (baris[KOLOM_SUB_PEKERJAAN]) subTerakhir = baris[KOLOM_SUB_PEKERJAAN];
      return {
        ...baris,
        [KOLOM_GEDUNG]: baris[KOLOM_GEDUNG] || gedungTerakhir,
        [KOLOM_SUB_PEKERJAAN]: baris[KOLOM_SUB_PEKERJAAN] || subTerakhir,
      };
    });

    const dataBersih = dataDiisi.filter(
      (b) => b[KOLOM_GEDUNG] && b[KOLOM_SUB_PEKERJAAN] && b[KOLOM_TITIK_LOKASI]
    );
    console.log(`✅ Berhasil memuat ${dataBersih.length} baris data titik lokasi dari ${DATA_FILE_PATH}`);
    return dataBersih;
  } catch (err) {
    console.error(`❌ Gagal membaca ${DATA_FILE_PATH}:`, err.message);
    return [];
  }
}

const daftarTitikLokasi = bacaDataTitikLokasi();

// Bangun daftar gedung final: urutan dari DAFTAR_GEDUNG_DASAR tetap dipertahankan,
// lalu gedung-gedung baru yang ditemukan di spreadsheet tapi belum ada di daftar
// dasar ditambahkan di akhir secara otomatis.
function bangunDaftarGedung() {
  const sudahAda = new Set(DAFTAR_GEDUNG_DASAR.map((g) => g.toLowerCase()));
  const tambahan = [];
  for (const baris of daftarTitikLokasi) {
    const nama = baris[KOLOM_GEDUNG];
    if (nama && !sudahAda.has(nama.toLowerCase())) {
      sudahAda.add(nama.toLowerCase());
      tambahan.push(nama);
    }
  }
  if (tambahan.length > 0) {
    console.log(`ℹ️  ${tambahan.length} gedung baru ditemukan di spreadsheet dan ditambahkan ke menu: ${tambahan.join(", ")}`);
  }
  return [...DAFTAR_GEDUNG_DASAR, ...tambahan];
}
const DAFTAR_GEDUNG = bangunDaftarGedung();

function waktuSekarang() {
  const sekarang = new Date();
  const tanggal = sekarang.toLocaleDateString("id-ID", { timeZone: "Asia/Jakarta", day: "2-digit", month: "2-digit", year: "numeric" }).replace(/\//g, "-");
  const bagianJam = new Intl.DateTimeFormat("en-GB", {
    timeZone: "Asia/Jakarta", hour: "2-digit", minute: "2-digit", hour12: false,
  }).formatToParts(sekarang);
  const jam = `${bagianJam.find((p) => p.type === "hour").value}:${bagianJam.find((p) => p.type === "minute").value}`;
  return { tanggal, jam };
}

function normalisasiTeks(t) {
  return String(t || "").trim().toLowerCase();
}

// ====== Sanitasi input untuk mencegah Excel formula injection ======
// Kalau teks dimulai dengan karakter yang bisa ditafsirkan Excel sebagai formula
// (=, +, -, @, \t, \r), awali dengan tanda kutip tunggal supaya Excel memperlakukannya
// sebagai teks biasa dan tidak mengeksekusi rumus apa pun.
function sanitasiUntukExcel(teks) {
  if (!teks) return teks;
  const KARAKTER_BERBAHAYA = /^[=+\-@\t\r]/;
  return KARAKTER_BERBAHAYA.test(teks) ? `'${teks}` : teks;
}

// ====== Rate limiting per pengirim (anti-flood) ======
const waktuPesanTerakhir = new Map(); // pengirim -> timestamp
function cekDanCatatRateLimit(pengirim) {
  const sekarang = Date.now();
  const terakhir = waktuPesanTerakhir.get(pengirim) || 0;
  if (sekarang - terakhir < COOLDOWN_MS) return false; // terlalu cepat, tolak
  waktuPesanTerakhir.set(pengirim, sekarang);
  return true; // boleh diproses
}

// ====== Resolusi nomor WA asli dari @lid (best-effort) ======
// WhatsApp kadang mengirim ID tersembunyi (@lid) alih-alih nomor asli demi privasi.
// Membalikkan @lid -> nomor asli TIDAK selalu bisa dijamin berhasil (ini keterbatasan
// dari sisi WhatsApp sendiri, bukan bug kode). Kalau Baileys kebetulan sudah tahu
// pemetaannya (dari pesan itu sendiri atau riwayat sebelumnya), kita pakai; kalau
// belum, tetap fallback aman ke ID lama supaya bot tidak error.
async function ambilNomorAsli(sock, pesan) {
  const remoteJid = pesan.key.remoteJid;
  if (!remoteJid || !remoteJid.endsWith("@lid")) return remoteJid; // sudah nomor asli, tidak perlu apa-apa

  if (pesan.key.remoteJidAlt && pesan.key.remoteJidAlt.endsWith("@s.whatsapp.net")) {
    return pesan.key.remoteJidAlt;
  }
  try {
    const pn = await sock.signalRepository?.lidMapping?.getPNForLID(remoteJid);
    if (pn) return pn;
  } catch {
    // abaikan, fallback ke bawah
  }
  return null; // belum bisa di-resolve
}

function formatNomorUntukLaporan(jidAsli, jidLid) {
  if (jidAsli) {
    const angkaSaja = jidAsli.split("@")[0].split(":")[0];
    return `+${angkaSaja}`;
  }
  const angkaLid = jidLid.split("@")[0].split(":")[0];
  return `${angkaLid} (ID tersembunyi, nomor asli belum diketahui)`;
}

// ====== Registrasi nomor & pengingat harian "belum lapor" ======
const NOMOR_TERDAFTAR_PATH = "nomor-terdaftar.json"; // { [pengirim]: { nama, terakhirLapor: "DD-MM-YYYY" } }
const JAM_PENGINGAT_HARIAN = 16; // jam 16:00 WIB
let tanggalPengingatTerakhir = null; // cegah kirim pengingat dobel di hari yang sama

async function muatNomorTerdaftar() {
  if (!fs.existsSync(NOMOR_TERDAFTAR_PATH)) return {};
  try {
    const isi = await fsPromises.readFile(NOMOR_TERDAFTAR_PATH, "utf8");
    return JSON.parse(isi);
  } catch {
    return {};
  }
}

async function catatLaporanHariIni(pengirim, namaPengirim) {
  const data = await muatNomorTerdaftar();
  const { tanggal } = waktuSekarang();
  data[pengirim] = { nama: namaPengirim || data[pengirim]?.nama || "Tidak diketahui", terakhirLapor: tanggal };
  await fsPromises.writeFile(NOMOR_TERDAFTAR_PATH, JSON.stringify(data, null, 2), "utf8");
}

async function cekDanKirimPengingatHarian(sock) {
  const jamSekarang = Number(
    new Intl.DateTimeFormat("en-GB", { timeZone: "Asia/Jakarta", hour: "2-digit", hour12: false }).format(new Date())
  );
  const { tanggal: tanggalHariIni } = waktuSekarang();

  if (jamSekarang < JAM_PENGINGAT_HARIAN) return; // belum waktunya
  if (tanggalPengingatTerakhir === tanggalHariIni) return; // sudah dikirim hari ini, jangan dobel

  const data = await muatNomorTerdaftar();
  for (const [pengirim, info] of Object.entries(data)) {
    if (info.terakhirLapor !== tanggalHariIni) {
      try {
        const sapaan = info.nama && info.nama !== "Tidak diketahui" ? ` ${info.nama}` : "";
        await sock.sendMessage(pengirim, {
          text: `Halo${sapaan} 👋\n\nHari ini (${tanggalHariIni}) belum ada laporan progres yang masuk dari nomor ini. Mohon segera isi laporan ya, ketik *menu* untuk mulai. 🙏`,
        });
        console.log(`   📨 Pengingat 'belum lapor hari ini' terkirim ke ${pengirim}`);
      } catch (err) {
        console.error(`   ❌ Gagal kirim pengingat harian ke ${pengirim}:`, err.message);
      }
      // Jeda 2-4 detik antar pesan supaya tidak kena deteksi spam WhatsApp
      await new Promise((r) => setTimeout(r, 2000 + Math.random() * 2000));
    }
  }
  tanggalPengingatTerakhir = tanggalHariIni;
}

function ambilTitikLokasi(gedung, subPekerjaan) {
  const gedungN = normalisasiTeks(gedung);
  const subN = normalisasiTeks(subPekerjaan);
  const hasil = daftarTitikLokasi.filter(
    (b) => normalisasiTeks(b[KOLOM_GEDUNG]) === gedungN && normalisasiTeks(b[KOLOM_SUB_PEKERJAAN]) === subN
  );
  // dedup nama titik lokasi yang sama
  const sudahAda = new Set();
  const unik = [];
  for (const item of hasil) {
    const nama = item[KOLOM_TITIK_LOKASI];
    if (!sudahAda.has(nama)) {
      sudahAda.add(nama);
      unik.push(item);
    }
  }
  return unik;
}

// ====== FUNGSI: Simpan laporan ke file CSV ======
// ====== FUNGSI: Simpan foto yang dikirim pekerja ke folder lokal ======
async function simpanFotoLaporan(sock, pesan, pengirim) {
  try {
    if (!fs.existsSync(FOLDER_FOTO)) fs.mkdirSync(FOLDER_FOTO);
    const buffer = await downloadMediaMessage(pesan, "buffer", {});
    const nomorBersih = pengirim.replace(/[^0-9]/g, "");
    const namaFile = `${Date.now()}_${nomorBersih}.jpg`;
    const filePath = path.join(FOLDER_FOTO, namaFile);
    fs.writeFileSync(filePath, buffer);
    return namaFile;
  } catch (err) {
    console.error("❌ Gagal menyimpan foto:", err.message);
    return null;
  }
}

const KOLOM_LAPORAN = [
  "Waktu Input", "Nomor Pengirim", "Nama Pengirim", "Gedung", "Tanggal Pengerjaan", "Jam Selesai",
  "Sub Pekerjaan", "Titik Lokasi", "Progres", "Status", "Kendala",
];

async function bukaAtauBuatWorkbookLaporan(filePath) {
  const workbook = new ExcelJS.Workbook();
  if (fs.existsSync(filePath)) {
    await workbook.xlsx.readFile(filePath);
  }
  let sheet = workbook.getWorksheet("Laporan");
  if (!sheet) {
    sheet = workbook.addWorksheet("Laporan");
    sheet.addRow([...KOLOM_LAPORAN, "Foto"]);
    sheet.getRow(1).font = { bold: true };
    sheet.columns = [...KOLOM_LAPORAN, "Foto"].map((judul) => ({
      width: judul === "Foto" ? 18 : judul === "Progres" || judul === "Titik Lokasi" ? 35 : 18,
    }));
  }
  return { workbook, sheet };
}

async function tulisSatuLaporanKeExcel(filePath, barisBaru, namaFileFoto) {
  const { workbook, sheet } = await bukaAtauBuatWorkbookLaporan(filePath);
  const nilaiBaris = KOLOM_LAPORAN.map((kolom) => barisBaru[kolom] ?? "");
  const rowBaru = sheet.addRow(nilaiBaris);
  const nomorBaris = rowBaru.number;

  if (namaFileFoto) {
    const fotoPath = path.join(FOLDER_FOTO, namaFileFoto);
    if (fs.existsSync(fotoPath)) {
      try {
        const imageId = workbook.addImage({ filename: fotoPath, extension: "jpeg" });
        sheet.addImage(imageId, {
          tl: { col: KOLOM_LAPORAN.length, row: nomorBaris - 1 },
          ext: { width: 100, height: 100 },
        });
        sheet.getRow(nomorBaris).height = 80;
      } catch (err) {
        console.error("⚠️  Gagal sisip foto ke Excel (tetap tercatat nama filenya):", err.message);
      }
    }
  }

  await workbook.xlsx.writeFile(filePath);
}

// ====== Google Sheets via Apps Script Web App (opsional) ======
async function kirimKeGoogleSheets(barisBaru, namaFileFoto) {
  if (!APPS_SCRIPT_URL) return false; // belum di-setup, skip diam-diam

  const payload = { token: APPS_SCRIPT_TOKEN, ...barisBaru };

  // Sertakan foto sebagai base64 supaya Apps Script bisa upload ke Drive & tampilkan gambarnya
  if (namaFileFoto) {
    const fotoPath = path.join(FOLDER_FOTO, namaFileFoto);
    if (fs.existsSync(fotoPath)) {
      try {
        const buffer = fs.readFileSync(fotoPath);
        payload.fotoBase64 = buffer.toString("base64");
        payload.fotoMime = "image/jpeg";
      } catch (err) {
        console.error("⚠️  Gagal baca file foto untuk dikirim ke Google Sheets:", err.message);
      }
    }
  }

  try {
    const kontroler = new AbortController();
    const timeoutId = setTimeout(() => kontroler.abort(), 30000); // 30 detik, foto butuh waktu lebih lama

    const respons = await fetch(APPS_SCRIPT_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
      signal: kontroler.signal,
    });
    clearTimeout(timeoutId);

    const hasil = await respons.json();
    if (hasil.ok) {
      console.log("✅ Laporan berhasil tersimpan ke Google Sheets.");
      return true;
    }

    console.error("❌ Google Sheets menolak data:", hasil.error);
    return false;
  } catch (err) {
    console.error("❌ Gagal kirim ke Google Sheets (mungkin internet putus):", err.message);
    return false;
  }
}

async function simpanLaporan(data) {
  const waktuInput = new Date().toLocaleString("id-ID", { timeZone: "Asia/Jakarta" });

  const barisBaru = {
    waktuInput,
    pengirim: data.pengirim,
    namaPengirim: data.namaPengirim || "Tidak diketahui",
    gedung: data.gedung,
    tanggalMulai: data.tanggalMulai,
    jamSelesai: data.jamSelesai,
    subPekerjaan: data.subPekerjaan,
    titikLokasi: data.titikLokasi,
    progres: data.progres,
    status: data.status,
    kendala: data.kendala,
  };
  // versi dengan nama kolom berspasi, dipakai khusus untuk nulis ke Excel lokal
  const barisBaruExcel = {
    "Waktu Input": waktuInput,
    "Nomor Pengirim": data.pengirim,
    "Nama Pengirim": data.namaPengirim || "Tidak diketahui",
    "Gedung": data.gedung,
    "Tanggal Pengerjaan": data.tanggalMulai,
    "Jam Selesai": data.jamSelesai,
    "Sub Pekerjaan": data.subPekerjaan,
    "Titik Lokasi": data.titikLokasi,
    "Progres": data.progres,
    "Status": data.status,
    "Kendala": data.kendala,
  };

  // Coba kirim ke Google Sheets dulu (kalau sudah di-setup)
  const berhasilKeGoogle = await kirimKeGoogleSheets(barisBaru, data.namaFileFoto);
  if (berhasilKeGoogle) {
    // Tetap simpan salinan ke Excel lokal juga sebagai cadangan, tapi tidak masalah kalau gagal
    tulisSatuLaporanKeExcel(LAPORAN_FILE_PATH, barisBaruExcel, data.namaFileFoto).catch(() => {});
    return { berhasil: true, tujuan: "google" };
  }

  // Fallback: simpan ke Excel lokal (foto ikut disisipkan langsung di sel)
  try {
    await tulisSatuLaporanKeExcel(LAPORAN_FILE_PATH, barisBaruExcel, data.namaFileFoto);
    return { berhasil: true, tujuan: "excel" };
  } catch (err) {
    console.error(`❌ Gagal menyimpan laporan ke ${LAPORAN_FILE_PATH}:`, err.message);
    console.error(`   Kemungkinan file sedang dibuka di Excel/aplikasi lain. Tutup dulu filenya.`);
    // Simpan ke file cadangan supaya laporan tidak hilang meski gagal tulis ke file utama
    try {
      await tulisSatuLaporanKeExcel(LAPORAN_CADANGAN_PATH, barisBaruExcel, data.namaFileFoto);
      console.log(`   Laporan disimpan sementara ke ${LAPORAN_CADANGAN_PATH}`);
    } catch {
      console.error(`   ❌ Gagal juga simpan ke file cadangan.`);
    }
    return { berhasil: false, tujuan: null };
  }
}

// ====== Session state per pengguna (disimpan juga ke file, tahan restart) ======
const SESI_FILE_PATH = "sesi-aktif.json";

function muatSesiDariFile() {
  if (!fs.existsSync(SESI_FILE_PATH)) return new Map();
  try {
    const isi = JSON.parse(fs.readFileSync(SESI_FILE_PATH, "utf8"));
    const map = new Map(Object.entries(isi));
    console.log(`✅ Memuat ${map.size} sesi aktif dari ${SESI_FILE_PATH} (melanjutkan laporan yang sempat terputus).`);
    return map;
  } catch {
    console.warn(`⚠️  Gagal membaca ${SESI_FILE_PATH}, mulai dari sesi kosong.`);
    return new Map();
  }
}

function simpanSemuaSesiKeFile() {
  const obj = Object.fromEntries(sesiPengguna);
  // Fire-and-forget async write -- in-memory Map tetap jadi sumber kebenaran,
  // file hanya untuk persistence antar restart.
  fsPromises.writeFile(SESI_FILE_PATH, JSON.stringify(obj, null, 2), "utf8").catch((err) => {
    console.error(`⚠️  Gagal menyimpan sesi ke ${SESI_FILE_PATH}:`, err.message);
  });
}

// Bungkus sesiPengguna.set supaya setiap perubahan otomatis tersimpan ke file
const sesiPengguna = muatSesiDariFile();
const _sesiSetAsli = sesiPengguna.set.bind(sesiPengguna);
sesiPengguna.set = (key, value) => {
  const hasil = _sesiSetAsli(key, value);
  simpanSemuaSesiKeFile();
  return hasil;
};

function teksMenuGedung() {
  const daftar = DAFTAR_GEDUNG.map((g, i) => `${i + 1}. ${g}`).join("\n");
  return `Selamat datang di CS Laporan Progress 👋\nSilakan pilih gedung (ketik nomornya):\n\n${daftar}\n\nKetik *menu* kapan saja untuk kembali ke sini, atau *batal* untuk membatalkan laporan yang sedang diisi.`;
}

function teksTitikLokasi(daftar) {
  return daftar.map((item, i) => `${i + 1}. ${item[KOLOM_TITIK_LOKASI]}`).join("\n");
}

// ====== FUNGSI UTAMA: Jalankan koneksi WhatsApp ======
// ====== FUNGSI: Kirim pesan WA dengan retry otomatis ======
// Dipakai supaya kalau pengiriman gagal (misal koneksi baru pulih dari putus
// sebentar dan belum benar-benar stabil), bot coba lagi otomatis beberapa kali
// alih-alih diam saja / pesan hilang tanpa balasan ke pekerja.
async function kirimPesanDenganRetry(sock, pengirim, isiPesan, percobaanMaksimal = 3) {
  for (let percobaan = 1; percobaan <= percobaanMaksimal; percobaan++) {
    try {
      await sock.sendMessage(pengirim, isiPesan);
      return true;
    } catch (err) {
      console.error(`⚠️  Gagal kirim pesan ke ${pengirim} (percobaan ${percobaan}/${percobaanMaksimal}):`, err.message);
      if (percobaan < percobaanMaksimal) {
        await new Promise((resolve) => setTimeout(resolve, 2000 * percobaan)); // jeda makin lama tiap percobaan
      }
    }
  }
  console.error(`❌ Gagal kirim pesan ke ${pengirim} setelah ${percobaanMaksimal} percobaan. Pesan tidak terkirim.`);
  return false;
}

// Versi Baileys di-cache di sini, cuma diambil sekali dari internet saat bot pertama
// kali start -- BUKAN diambil ulang tiap kali reconnect. Ini penting supaya reconnect
// setelah laptop sleep tidak ikut menunggu request internet (fetchLatestBaileysVersion)
// yang bisa lambat/nge-hang kalau WiFi belum benar-benar stabil lagi.
let versiBaileysCache = null;

let sockAktif = null; // referensi koneksi yang sedang hidup, supaya bisa ditutup rapi sebelum bikin yang baru
let sedangReconnect = false; // cegah 2 percobaan reconnect jalan bersamaan

async function mulaiBot() {
  // Kalau ada koneksi lama yang masih menggantung (belum benar-benar tertutup),
  // tutup dulu secara eksplisit sebelum bikin koneksi baru. Tanpa ini, dalam
  // kondisi tertentu (misalnya reconnect cepat berkali-kali setelah laptop sleep)
  // bisa ada 2 koneksi hidup bersamaan yang rebutan sesi WhatsApp yang sama --
  // inilah kemungkinan besar penyebab error 'conflict' berulang & bot crash total.
  if (sockAktif) {
    try {
      sockAktif.ev.removeAllListeners();
      sockAktif.end(undefined);
    } catch (err) {
      console.warn("⚠️  Gagal menutup koneksi lama dengan rapi (dilanjutkan saja):", err.message);
    }
    sockAktif = null;
  }

  const { state, saveCreds } = await useMultiFileAuthState("auth_info");
  if (!versiBaileysCache) {
    const { version } = await fetchLatestBaileysVersion();
    versiBaileysCache = version;
  }

  const sock = makeWASocket({
    version: versiBaileysCache,
    auth: state,
    logger: pino({ level: "silent" }),
    printQRInTerminal: false,
    // Cek koneksi tiap 10 detik (lebih cepat dari default) supaya kalau laptop
    // sleep/internet putus, bot lebih cepat SADAR koneksinya mati dan langsung
    // reconnect -- bukan baru sadar setelah beberapa menit.
    keepAliveIntervalMs: 10_000,
  });
  sockAktif = sock;

  sock.ev.on("creds.update", saveCreds);

  sock.ev.on("connection.update", async (update) => {
    const { connection, lastDisconnect, qr } = update;
    if (qr) {
      console.log("\n📱 Scan QR code ini dengan WhatsApp (Perangkat Tertaut):\n");
      qrcode.generate(qr, { small: true });
    }
    if (connection === "close") {
      const alasan = lastDisconnect?.error?.output?.statusCode;
      const pesanAlasan = lastDisconnect?.error?.message || "(tidak ada pesan)";
      const harusReconnect = alasan !== DisconnectReason.loggedOut;
      console.log(`Koneksi terputus. Kode alasan: ${alasan} | Pesan: ${pesanAlasan} | Reconnect: ${harusReconnect}`);
      if (harusReconnect) {
        // Jeda sebentar sebelum coba reconnect, supaya tidak spam percobaan kalau
        // internet/WA sedang benar-benar bermasalah. Errornya juga ditangkap (.catch)
        // supaya kalau reconnect gagal, program TIDAK crash total -- cukup log saja
        // dan biarkan event "close" berikutnya (kalau ada) yang coba lagi.
        // Guard sedangReconnect: cegah 2 percobaan reconnect jalan bersamaan kalau
        // event "close" sempat terpicu berkali-kali dengan cepat (misalnya pas laptop
        // baru bangun dari sleep dan koneksi masih naik-turun sebentar).
        if (!sedangReconnect) {
          sedangReconnect = true;
          setTimeout(() => {
            mulaiBot()
              .catch((err) => {
                console.error("❌ Gagal reconnect:", err.message);
              })
              .finally(() => {
                sedangReconnect = false;
              });
          }, 3000);
        }
      }
    } else if (connection === "open") {
      console.log("✅ Bot CS WhatsApp berhasil terhubung dan siap menerima pesan!");

      // Beri tahu pekerja yang laporannya sempat tertunda (belum selesai) kalau bot sudah aktif lagi.
      // Ditandai per-sesi (bukan sekali untuk selamanya) supaya: (1) tiap kali bot benar-benar
      // reconnect setelah error, sesi yang belum pernah diingatkan tetap dapat notifikasi baru,
      // tapi (2) sesi yang SAMA tidak diingatkan berulang-ulang kalau error lagi sementara pekerja
      // itu masih belum lanjut isi (anti-spam). Begitu pekerja mulai laporan baru, sesinya diganti
      // objek baru (tanpa flag ini) sehingga otomatis "reset" untuk laporan berikutnya.
      for (const [pengirim, sesi] of sesiPengguna.entries()) {
        if (sesi.step && sesi.step !== "pilih_gedung" && !sesi.sudahDiingatkanBotAktif) {
          try {
            await kirimPesanDenganRetry(sock, pengirim, {
              text: "Bot sudah aktif kembali ✅\n\nLaporan Anda yang belum selesai masih tersimpan, silakan lanjutkan mengisi seperti biasa. Ketik *menu* kalau mau mulai ulang dari awal.",
            });
            console.log(`   📨 Notifikasi 'bot aktif lagi' terkirim ke ${pengirim}`);
          } catch (err) {
            console.error(`   ❌ Gagal kirim notifikasi ke ${pengirim}:`, err.message);
          }
          // Tandai sudah diingatkan (pakai .set supaya otomatis tersimpan ke sesi-aktif.json juga)
          sesiPengguna.set(pengirim, { ...sesi, sudahDiingatkanBotAktif: true });
        }
      }
    }
  });

  sock.ev.on("messages.upsert", async ({ messages, type }) => {
    // Log diagnostik: catat SEMUA event upsert yang masuk (termasuk yang nanti akan
    // diabaikan), supaya kalau ada pesan yang "hilang" lagi, kita bisa lihat persis
    // jenis event & alasan sebenarnya dari log, bukan tebak-tebakan.
    const pesanMentah = messages?.[0];
    console.log(`[DEBUG upsert] type=${type} | dariSaya=${pesanMentah?.key?.fromMe} | adaIsiPesan=${!!pesanMentah?.message} | timestamp=${pesanMentah?.messageTimestamp}`);

    // "notify" = pesan baru realtime. "append" = pesan susulan yang tersinkron
    // saat bot baru online lagi (misal setelah offline/restart) -- ini yang bikin
    // pesan yang terkirim waktu bot mati tetap ikut diproses & dibalas.
    if (type !== "notify" && type !== "append") return;

    const pesan = messages[0];
    if (!pesan.message || pesan.key.fromMe) return;

    // Guard: jangan proses pesan yang sudah SANGAT lama (>12 jam), supaya tidak
    // tiba-tiba membalas riwayat chat lama kalau bot mati berhari-hari. Tapi cukup
    // longgar supaya pesan yang masuk waktu laptop sempat sleep/internet putus
    // beberapa jam (istirahat, semalaman, dll) tetap ikut diproses begitu bot nyala lagi.
    const BATAS_USIA_PESAN_MS = 12 * 60 * 60 * 1000; // 12 jam
    const timestampPesan = Number(pesan.messageTimestamp) * 1000;
    if (timestampPesan && Date.now() - timestampPesan > BATAS_USIA_PESAN_MS) return;

    const pengirim = pesan.key.remoteJid;
    const namaPengirim = pesan.pushName || "Tidak diketahui";
    const isGroup = pengirim?.endsWith("@g.us");
    const isChannel = pengirim?.endsWith("@newsletter");
    if (isGroup || isChannel) return;

    // Coba resolusi nomor WA asli (kalau pengirim pakai @lid) -- HANYA untuk keperluan
    // tampilan/laporan, TIDAK dipakai untuk membalas pesan (tetap pakai `pengirim` asli
    // apa adanya supaya balasan tetap sampai ke chat yang benar).
    const nomorAsli = pengirim?.endsWith("@lid") ? await ambilNomorAsli(sock, pesan) : pengirim;
    const nomorUntukLaporan = formatNomorUntukLaporan(nomorAsli, pengirim);

    const adaGambar = !!pesan.message.imageMessage;
    const teksMasuk =
      pesan.message.conversation ||
      pesan.message.extendedTextMessage?.text ||
      pesan.message.imageMessage?.caption ||
      "";
    if (!teksMasuk && !adaGambar) return;

    console.log(`📩 Pesan masuk dari ${pengirim} (${type})${adaGambar ? " [ada foto]" : ""}: ${teksMasuk}`);

    // Rate limit: abaikan pesan kalau pengirim yang sama mengirim terlalu cepat
    if (!cekDanCatatRateLimit(pengirim)) return;

    await sock.sendPresenceUpdate("composing", pengirim);

    const teksAsli = teksMasuk.trim();
    const teksBersih = teksAsli.toLowerCase();
    let sesi = sesiPengguna.get(pengirim);

    if (teksBersih === "batal" && sesi) {
      sesiPengguna.set(pengirim, { step: "pilih_gedung", laporanTerakhir: sesi.laporanTerakhir });
      await kirimPesanDenganRetry(sock, pengirim, { text: `Laporan dibatalkan.\n\n${teksMenuGedung()}` });
      return;
    }

    if (teksBersih === "ulang") {
      if (!sesi?.laporanTerakhir) {
        await kirimPesanDenganRetry(sock, pengirim, {
          text: `Belum ada laporan sebelumnya yang bisa diulang. Ketik *menu* untuk mulai lapor dari awal.`,
        });
        return;
      }
      const { gedung, subPekerjaan, titikLokasi } = sesi.laporanTerakhir;
      const { tanggal, jam } = waktuSekarang();
      sesiPengguna.set(pengirim, {
        step: "input_progres",
        gedung, subPekerjaan, titikLokasi,
        tanggalMulai: tanggal, jamSelesai: jam, namaPengirim,
        nomorUntukLaporan: sesi.nomorUntukLaporan || nomorUntukLaporan,
        laporanTerakhir: sesi.laporanTerakhir,
      });
      await kirimPesanDenganRetry(sock, pengirim, {
        text:
          `Melanjutkan laporan di lokasi yang sama:\n\n` +
          `Gedung: *${gedung}*\nSub Pekerjaan: *${subPekerjaan}*\nTitik Lokasi: *${titikLokasi}*\n` +
          `Tanggal: *${tanggal}* | Jam: *${jam}* (otomatis)\n\n` +
          `Ketik progres pekerjaan terbaru untuk lokasi ini.`,
      });
      return;
    }

    if (teksBersih === "menu" || !sesi) {
      sesiPengguna.set(pengirim, { step: "pilih_gedung", laporanTerakhir: sesi?.laporanTerakhir });
      await kirimPesanDenganRetry(sock, pengirim, { text: teksMenuGedung() });
      return;
    }

    // ====== HALAMAN 1: Pilih gedung ======
    if (sesi.step === "pilih_gedung") {
      const inputAngkaMurni = /^\d+$/.test(teksAsli);
      const nomor = parseInt(teksBersih, 10);
      if (!inputAngkaMurni || !nomor || nomor < 1 || nomor > DAFTAR_GEDUNG.length) {
        // Kalau yang diketik bukan angka sama sekali (misalnya sapaan "hai"),
        // jangan tampilkan "Nomor tidak valid" -- cukup tampilkan menu apa adanya.
        const awalan = inputAngkaMurni ? "Nomor tidak valid.\n\n" : "";
        await kirimPesanDenganRetry(sock, pengirim, { text: `${awalan}${teksMenuGedung()}` });
        return;
      }
      const gedung = DAFTAR_GEDUNG[nomor - 1];
      const { tanggal, jam } = waktuSekarang();
      sesiPengguna.set(pengirim, { step: "pilih_sub_pekerjaan", gedung, tanggalMulai: tanggal, jamSelesai: jam, namaPengirim, nomorUntukLaporan, laporanTerakhir: sesi.laporanTerakhir });
      await kirimPesanDenganRetry(sock, pengirim, {
        text: `Gedung: *${gedung}*\nTanggal: *${tanggal}* | Jam: *${jam}* (otomatis)\n\nPilih sub pekerjaan:\n1️⃣ Mekanikal\n2️⃣ Elektrikal`,
      });
      return;
    }

    // ====== STEP 4: Pilih sub pekerjaan ======
    // ====== HALAMAN 3: Pilih sub pekerjaan ======
    if (sesi.step === "pilih_sub_pekerjaan") {
      if (teksBersih !== "1" && teksBersih !== "2") {
        await kirimPesanDenganRetry(sock, pengirim, { text: "Mohon ketik 1 (Mekanikal) atau 2 (Elektrikal)." });
        return;
      }
      const subPekerjaan = teksBersih === "1" ? "Mekanikal" : "Elektrikal";
      const daftarTitik = ambilTitikLokasi(sesi.gedung, subPekerjaan);

      if (daftarTitik.length === 0) {
        await kirimPesanDenganRetry(sock, pengirim, {
          text: `Belum ada data Titik Lokasi untuk *${sesi.gedung}* - *${subPekerjaan}*. Ketik *menu* untuk mulai ulang.`,
        });
        return;
      }

      sesiPengguna.set(pengirim, { ...sesi, step: "pilih_titik_lokasi", subPekerjaan, daftarTitik });
      await kirimPesanDenganRetry(sock, pengirim, {
        text: `Sub Pekerjaan: *${subPekerjaan}*\n\nPilih Titik Lokasi:\n\n${teksTitikLokasi(daftarTitik)}\n\nKetik nomornya.`,
      });
      return;
    }

    // ====== Pilih titik lokasi ======
    if (sesi.step === "pilih_titik_lokasi") {
      const nomor = parseInt(teksBersih, 10);
      if (!nomor || nomor < 1 || nomor > sesi.daftarTitik.length) {
        await kirimPesanDenganRetry(sock, pengirim, {
          text: `Nomor tidak valid. Pilih 1-${sesi.daftarTitik.length}.`,
        });
        return;
      }
      const titikLokasi = sesi.daftarTitik[nomor - 1][KOLOM_TITIK_LOKASI];
      sesiPengguna.set(pengirim, { ...sesi, step: "input_progres", titikLokasi });
      await kirimPesanDenganRetry(sock, pengirim, {
        text:
          `Titik Lokasi: *${titikLokasi}*\n\n` +
          `Ketik progres pekerjaan: jelaskan pekerjaan yang dilakukan dan sejauh mana progresnya (jumlah titik/panjang/persentase, sesuai jenis pekerjaannya).\n\n` +
          `Contoh:\n` +
          `• Pemasangan titik lampu (3/10) titik\n` +
          `• Instalasi kabel feeder (50/120) meter\n` +
          `• Penarikan kabel MDP-2.1 ke SDP-8 (60%)\n` +
          `• Pemasangan panel selesai, tinggal wiring\n\n` +
          `Boleh disesuaikan dengan progres pekerjaan yang sebenarnya.`,
      });
      return;
    }

    // ====== Input progres ======
    if (sesi.step === "input_progres") {
      if (teksAsli.length > BATAS_PANJANG_INPUT) {
        await kirimPesanDenganRetry(sock, pengirim, {
          text: `Teks terlalu panjang (maks ${BATAS_PANJANG_INPUT} karakter). Mohon ringkas dan kirim ulang.`,
        });
        return;
      }
      sesiPengguna.set(pengirim, { ...sesi, step: "input_foto", progres: teksAsli });
      await kirimPesanDenganRetry(sock, pengirim, {
        text: `Progres tercatat.\n\nMau lampirkan foto? Kirim foto sekarang, atau ketik "-" untuk lewati.`,
      });
      return;
    }

    // ====== Input foto (opsional) ======
    if (sesi.step === "input_foto") {
      if (adaGambar) {
        const namaFileFoto = await simpanFotoLaporan(sock, pesan, pengirim);
        sesiPengguna.set(pengirim, { ...sesi, step: "pilih_status", namaFileFoto: namaFileFoto || "" });
        await kirimPesanDenganRetry(sock, pengirim, {
          text: namaFileFoto
            ? `📷 Foto tersimpan.\n\nPilih status:\n1️⃣ Open\n2️⃣ Close`
            : `⚠️ Gagal menyimpan foto, tapi laporan tetap lanjut tanpa foto.\n\nPilih status:\n1️⃣ Open\n2️⃣ Close`,
        });
        return;
      }
      if (teksBersih === "-") {
        sesiPengguna.set(pengirim, { ...sesi, step: "pilih_status", namaFileFoto: "" });
        await kirimPesanDenganRetry(sock, pengirim, { text: `Dilewati (tanpa foto).\n\nPilih status:\n1️⃣ Open\n2️⃣ Close` });
        return;
      }
      await kirimPesanDenganRetry(sock, pengirim, {
        text: `Mohon kirim foto, atau ketik "-" untuk lewati tanpa foto.`,
      });
      return;
    }

    // ====== Pilih status ======
    if (sesi.step === "pilih_status") {
      if (teksBersih !== "1" && teksBersih !== "2") {
        await kirimPesanDenganRetry(sock, pengirim, { text: "Mohon ketik 1 (Open) atau 2 (Close)." });
        return;
      }
      const status = teksBersih === "1" ? "Open" : "Close";
      sesiPengguna.set(pengirim, { ...sesi, step: "input_kendala", status });
      await kirimPesanDenganRetry(sock, pengirim, {
        text: `Status: *${status}*\n\nKetik kendala (jika ada), atau ketik "-" kalau tidak ada kendala.`,
      });
      return;
    }

    // ====== Input kendala -> simpan laporan ======
    if (sesi.step === "input_kendala") {
      if (teksAsli !== "-" && teksAsli.length > BATAS_PANJANG_INPUT) {
        await kirimPesanDenganRetry(sock, pengirim, {
          text: `Teks terlalu panjang (maks ${BATAS_PANJANG_INPUT} karakter). Mohon ringkas dan kirim ulang.`,
        });
        return;
      }
      const kendalaBersih = teksAsli === "-" ? "" : teksAsli;

      const hasilSimpan = await simpanLaporan({
        pengirim: sesi.nomorUntukLaporan || pengirim,
        namaPengirim: sanitasiUntukExcel(sesi.namaPengirim),
        gedung: sesi.gedung,
        tanggalMulai: sesi.tanggalMulai,
        jamSelesai: sesi.jamSelesai,
        subPekerjaan: sesi.subPekerjaan,
        titikLokasi: sesi.titikLokasi,
        progres: sanitasiUntukExcel(sesi.progres),
        status: sesi.status,
        kendala: sanitasiUntukExcel(kendalaBersih),
        namaFileFoto: sesi.namaFileFoto || "",
      });
      // Catat ke registry supaya sistem pengingat harian tahu nomor ini sudah lapor hari ini
      catatLaporanHariIni(pengirim, sesi.namaPengirim).catch((err) => {
        console.error("⚠️  Gagal catat laporan harian:", err.message);
      });

      const judulPesan = hasilSimpan.berhasil
        ? "✅ Laporan tersimpan!"
        : "⚠️ Laporan tercatat, tapi ada kendala teknis menyimpan ke file utama (kemungkinan file sedang dibuka admin). Data sudah aman di file cadangan, akan digabung nanti.";

      await kirimPesanDenganRetry(sock, pengirim, {
        text:
          `${judulPesan}\n\n` +
          `Gedung: ${sesi.gedung}\n` +
          `Tanggal: ${sesi.tanggalMulai} | Jam selesai: ${sesi.jamSelesai}\n` +
          `Sub Pekerjaan: ${sesi.subPekerjaan}\n` +
          `Titik Lokasi: ${sesi.titikLokasi}\n` +
          `Progres: ${sesi.progres}\n` +
          `Foto: ${sesi.namaFileFoto ? "✅ terlampir" : "(tidak ada)"}\n` +
          `Status: ${sesi.status}\n` +
          `Kendala: ${kendalaBersih || "(tidak ada)"}\n\n` +
          `Ketik *menu* untuk lapor pekerjaan lain, atau ketik *ulang* untuk lanjutkan update progres di lokasi yang sama tanpa pilih gedung dari awal.`,
      });
      console.log(`📝 Laporan disimpan ke: ${hasilSimpan.tujuan === "google" ? "Google Sheets" : hasilSimpan.tujuan === "excel" ? "Excel lokal" : "GAGAL semua"} | ${sesi.gedung} / ${sesi.titikLokasi} - ${sesi.status} (dari ${pengirim})`);
      sesiPengguna.set(pengirim, {
        step: "pilih_gedung",
        laporanTerakhir: { gedung: sesi.gedung, subPekerjaan: sesi.subPekerjaan, titikLokasi: sesi.titikLokasi },
      });
      return;
    }
  });
}

// ====== Pengaman tambahan: tangkap error yang tidak sengaja lolos dari try/catch ======
// manapun di kode, supaya SELALU ada log yang jelas kenapa proses berhenti/error,
// bukan cuma diam-diam mati tanpa keterangan (memudahkan diagnosa "putus-nyambung").
process.on("unhandledRejection", (err) => {
  console.error("❌ Unhandled Rejection (error async yang tidak tertangkap):", err);
});
process.on("uncaughtException", (err) => {
  console.error("❌ Uncaught Exception (error tidak tertangkap):", err);
});

// ====== Graceful shutdown: matikan bot dengan aman saat Ctrl+C / SIGTERM ======
// Menyimpan semua sesi aktif ke file dan menutup koneksi WhatsApp dengan rapi,
// supaya saat bot dinyalakan lagi, sesi pekerja yang sedang mengisi tidak hilang.
function matikanBotDenganAman(sinyal) {
  console.log(`\n🛑 Menerima sinyal ${sinyal} — mematikan bot dengan aman...`);
  try {
    // Simpan sesi terakhir ke file (synchronous karena proses akan segera keluar)
    const obj = Object.fromEntries(sesiPengguna);
    fs.writeFileSync(SESI_FILE_PATH, JSON.stringify(obj, null, 2), "utf8");
    console.log(`   ✅ ${sesiPengguna.size} sesi aktif tersimpan.`);
  } catch (err) {
    console.error("   ❌ Gagal menyimpan sesi:", err.message);
  }
  try {
    if (sockAktif) {
      sockAktif.ev.removeAllListeners();
      sockAktif.end(undefined);
      console.log("   ✅ Koneksi WhatsApp ditutup.");
    }
  } catch (err) {
    console.error("   ⚠️  Gagal menutup koneksi:", err.message);
  }
  // Tutup log stream sebelum keluar
  logStream.end(() => process.exit(0));
}
process.on("SIGINT", () => matikanBotDenganAman("SIGINT"));
process.on("SIGTERM", () => matikanBotDenganAman("SIGTERM"));

// ====== Cek berkala untuk kirim pengingat harian "belum lapor" ======
// Dicek tiap 15 menit; fungsi cekDanKirimPengingatHarian sendiri yang menentukan
// apakah sudah waktunya (>= JAM_PENGINGAT_HARIAN) dan belum terkirim hari ini.
setInterval(() => {
  if (sockAktif) {
    cekDanKirimPengingatHarian(sockAktif).catch((err) => {
      console.error("❌ Gagal cek/kirim pengingat harian:", err.message);
    });
  }
}, 15 * 60 * 1000);

mulaiBot().catch((err) => {
  console.error("❌ Gagal menjalankan bot saat start awal:", err.message);
  console.error("   Kalau ini terjadi berulang, cek koneksi internet PC/laptop ini.");
  process.exit(1); // keluar dengan kode error, supaya PM2 tahu harus otomatis restart proses
});
