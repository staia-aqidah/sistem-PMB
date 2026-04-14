// ═══════════════════════════════════════════════════════════════
// KONFIGURASI — ISI SEMUA ID DI BAGIAN INI SEBELUM DEPLOY
// ═══════════════════════════════════════════════════════════════

const SHEET_ID        = "1SZY9KfEXbidJJUCSns2lxNZuCzBjteHBofU1RRTbzQ8"; // Spreadsheet PMB
const FOLDER_ID       = "1liTy5F-qZSlHd5jah1Kd2OYWaCdgabuF";              // Folder upload berkas Drive
const DB_SHEET_ID     = "1pNHS6j8s2U5s5sBLqOOnGy6rQZ6K-nQ4HYsgAysMlOQ";
const KTM_TEMPLATE_ID = "1IFOkp_u76T3teu4zVeOQIl71dNRhUmVbu5nx_SJfbfM";
const KTM_FOLDER_ID   = "1t44UX7-LFfqg0PislheZVEywSc0e8H_J";

// ID folder untuk manajemen berkas mahasiswa yang sudah terverifikasi
// Buat folder "Berkas Mahasiswa Terverifikasi" di dalam "Database Mahasiswa"
const BERKAS_DB_FOLDER_ID = "1825Y-p-wtys49TuXEhzTFP4FBXBKvpJw";

// FOLDER_ID_PMB_BARU  = sub-folder "Mahasiswa Baru"  di dalam folder PMB
// FOLDER_ID_PMB_PINDAHAN = sub-folder "Mahasiswa Pindahan"
// Ini adalah folder yang sudah ada di sistem PMB Anda (FOLDER_ID sudah ada di atas
// adalah folder ROOT PMB, bukan subfolder. Tidak perlu diubah — kode akan
// mencari subfolder by name secara otomatis).


const KODE_PRODI_NIM = {
  KPI:110, PGMI:200, PAI:210, PIAUD:230, AS:310, HES:320
};

const NAMA_PRODI_LENGKAP = {
  KPI:   "Komunikasi dan Penyiaran Islam",
  PGMI:  "Pendidikan Guru Madrasah Ibtidaiyah",
  PAI:   "Pendidikan Agama Islam",
  PIAUD: "Pendidikan Islam Anak Usia Dini",
  AS:    "Ahwal Syakhshiyah",
  HES:   "Hukum Ekonomi Syariah"
};

const PRODI_DB_ID = {
  KPI:1, PGMI:2, PAI:3, PIAUD:4, AS:5, HES:6
};

// Sesuaikan dengan 2 fakultas kampus Anda
// Contoh: 1 = Tarbiyah  |  2 = Syariah & Dakwah
const FAKULTAS_DB_ID = {
  PAI:1, PGMI:1, PIAUD:1,
  KPI:2, AS:2,   HES:2
};

// Nama lengkap fakultas untuk KTM — sesuaikan dengan nama resmi kampus Anda
const NAMA_FAKULTAS = {
  1: "Tarbiyah",
  2: "Syariah dan Dakwah"
};

// Peta indeks kolom di sheet master_pmb (0-based, sesuai appendRow simpanData)
const C = {
  NO_URUT:0,  TIMESTAMP:1,  NIK:2,       NAMA:3,
  TEMPAT:4,   TGL_LHR:5,    JK:6,        NISN:7,
  SEKOLAH:8,  PRODI:9,      JENJANG:10,  KELAS:11,
  STAT_MHS:12, PT_ASAL:13,  JNJ_ASAL:14, PRODI_ASAL:15,
  ALAMAT:16,  EMAIL:17,     NO_HP:18,
  AYAH:19,    PKJ_AYAH:20,  IBU:21,      PKJ_IBU:22,  ALAMAT_ORTU:23,
  IJAZAH:24,  KTP:25,       KK:26,
  F23:27,     F34:28,       F46:29,
  KHS:30,     SURAT:31,
  REFERRAL:32,
  STATUS:33,   // Pending | Terverifikasi | Ditolak  (kolom 34 di Sheets, 1-based)
  CATATAN:34,  // kolom 35 (1-based)
  NIM:35       // kolom 36 (1-based) — diisi sistem saat Terverifikasi
};

// ═══════════════════════════════════════════════════════════════
// FORM PMB
// ═══════════════════════════════════════════════════════════════

function doGet(e) {
  const tpl = HtmlService.createTemplateFromFile("Form");
  tpl.cacheBust = new Date().getTime();
  tpl.refCode = (e && e.parameter && e.parameter.ref)
    ? e.parameter.ref.toString().toUpperCase().replace(/[^A-Z0-9\-_]/g,"").substring(0,20)
    : "";
  return tpl.evaluate()
    .setTitle("PMB 2025/2026 \u2014 IAI Al-Aqidah Al-Hasyimiyyah Jakarta")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function cekDuplikat(nik, nisn) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("master_pmb");
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][C.NIK]).trim()  === nik ||
          String(data[i][C.NISN]).trim() === nisn)
        return { duplikat:true, baris:i+1, nama:String(data[i][C.NAMA]).trim() };
    }
    return { duplikat:false };
  } catch(e) { return { duplikat:false }; }
}

function cekDuplikatForm(nik, nisn) {
  return cekDuplikat(nik.trim(), nisn.trim());
}

function uploadSingleFile(b64, fileName, mimeType, subFolderName, personalFolderName) {
  try {
    const root   = DriveApp.getFolderById(FOLDER_ID);
    const sub    = getOrCreateFolder(root, subFolderName);
    const person = getOrCreateFolder(sub, personalFolderName);
    const blob   = Utilities.newBlob(Utilities.base64Decode(b64), mimeType, fileName);
    const file   = person.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { success:true, url:file.getUrl() };
  } catch(e) { return { success:false, message:e.message }; }
}

function simpanData(fd) {
  try {
    const cek = cekDuplikat(fd.nik.trim(), fd.nisn.trim());
    if (cek.duplikat) return {
      success:false, duplikat:true,
      message:"NIK/NISN sudah terdaftar (baris "+cek.baris+", "+cek.nama+"). Hubungi panitia."
    };

    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("master_pmb");
    const no    = sheet.getLastRow();
    const ts    = new Date();

    sheet.appendRow([
      no, ts,                          // 0,1
      fd.nik,                          // 2  NIK
      fd.namaLengkap,                  // 3  NAMA
      fd.tempatLahir,                  // 4
      fd.tanggalLahir,                 // 5
      fd.jenisKelamin,                 // 6
      fd.nisn,                         // 7  NISN
      fd.sekolahAsal,                  // 8
      fd.prodiPilihan,                 // 9
      fd.jenjang        || "S1",       // 10
      fd.kelas,                        // 11
      fd.statusMahasiswa,              // 12
      fd.ptAsal         || "",         // 13
      fd.jenjangAsal    || "",         // 14
      fd.prodiAsal      || "",         // 15
      fd.alamatDomisili,               // 16
      fd.email,                        // 17
      fd.noHp           || "",         // 18
      fd.namaBapak,                    // 19
      fd.pekerjaanBapak,               // 20
      fd.namaIbu,                      // 21
      fd.pekerjaanIbu,                 // 22
      fd.alamatOrtu,                   // 23
      fd.urlIjazah,                    // 24
      fd.urlKTP,                       // 25
      fd.urlKK,                        // 26
      fd.urlFoto23,                    // 27
      fd.urlFoto34,                    // 28
      fd.urlFoto46,                    // 29
      fd.urlKHS         || "",         // 30
      fd.urlSurat       || "",         // 31
      fd.kodeReferral   || "",         // 32
      "Pending",                       // 33 STATUS
      ""                               // 34 CATATAN (kolom 36 = NIM diisi saat Terverifikasi)
    ]);

    return { success:true, noPendaftaran:"PMB-2025-"+String(no).padStart(4,"0") };
  } catch(e) { return { success:false, message:e.message }; }
}

function getOrCreateFolder(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

// ═══════════════════════════════════════════════════════════════
// REFERRAL
// ═══════════════════════════════════════════════════════════════

function cekReferral(kode) {
  if (!kode) return { valid:false, nama:"" };
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("referral");
    if (!sheet) return { valid:false, nama:"" };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toUpperCase() === kode.toUpperCase() && data[i][2] === true)
        return { valid:true, nama:data[i][1] };
    }
    return { valid:false, nama:"" };
  } catch(e) { return { valid:false, nama:"" }; }
}

function rekapReferral() {
  const masterSheet   = SpreadsheetApp.openById(SHEET_ID).getSheetByName("master_pmb");
  const referralSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("referral");
  if (!masterSheet || !referralSheet) return;

  const rows  = masterSheet.getDataRange().getValues();
  const rekap = {};
  for (let i = 1; i < rows.length; i++) {
    const kode   = String(rows[i][C.REFERRAL]).trim().toUpperCase();
    const status = String(rows[i][C.STATUS]).trim();
    if (!kode) continue;
    if (!rekap[kode]) rekap[kode] = { total:0, terverifikasi:0 };
    rekap[kode].total++;
    if (status === "Terverifikasi") rekap[kode].terverifikasi++;
  }

  const refRows = referralSheet.getDataRange().getValues();
  for (let i = 1; i < refRows.length; i++) {
    const kode = String(refRows[i][0]).trim().toUpperCase();
    const r    = rekap[kode] || { total:0, terverifikasi:0 };
    referralSheet.getRange(i+1, 4).setValue(r.total);
    referralSheet.getRange(i+1, 5).setValue(r.terverifikasi);
  }
}

// ═══════════════════════════════════════════════════════════════
// TRIGGER VERIFIKASI
// Pasang sebagai Installable Trigger:
//   Apps Script Editor → ⏰ Triggers → + Add Trigger
//   Function: onEditPMB | Event source: From spreadsheet | Event type: On edit
// ═══════════════════════════════════════════════════════════════

function onEditPMB(e) {
  try {
    const sheet = e.range.getSheet();
    if (sheet.getName() !== "master_pmb") return;
    if (e.range.getColumn() !== C.STATUS + 1) return; // kolom 34 (1-based)
    if (String(e.range.getValue()).trim() !== "Terverifikasi") return;

    const row = e.range.getRow();
    if (row <= 1) return;

    // Guard: skip jika NIM sudah ada (sudah diproses sebelumnya)
    if (sheet.getRange(row, C.NIM + 1).getValue()) return;

    prosesVerifikasi(sheet, row);

  } catch(err) {
    try {
      e.range.getSheet()
        .getRange(e.range.getRow(), C.CATATAN + 1)
        .setValue("⚠️ Error sistem: " + err.message);
    } catch(_) {}
    throw err;
  }
}

function prosesVerifikasi(sheet, row) {
  const data  = sheet.getRange(row, 1, 1, C.NIM + 1).getValues()[0];
  const prodi = String(data[C.PRODI]).trim().toUpperCase();

  if (!KODE_PRODI_NIM[prodi]) {
    sheet.getRange(row, C.CATATAN + 1)
         .setValue("⚠️ Kode prodi tidak dikenal: \"" + prodi + "\"");
    return;
  }

  const tahun  = new Date().getFullYear().toString().slice(-2);
  const noUrut = getNoUrutProdi(prodi, tahun);
  const nim    = tahun + String(KODE_PRODI_NIM[prodi]) + String(noUrut).padStart(4,"0");

  // Tulis NIM
  sheet.getRange(row, C.NIM + 1).setValue(nim);

  // Input ke database
  try {
    inputKeMahasiswaDB(nim, data, tahun);
  } catch(dbErr) {
    sheet.getRange(row, C.CATATAN + 1)
         .setValue("✅ NIM: " + nim + " | ⚠️ DB gagal: " + dbErr.message);
    return;
  }

  // Kelola folder berkas (shortcut di Database Mahasiswa)
  let berkasInfo = "";
  try {
    const bResult = kelolaFolderBerkas(nim, data, tahun);
    berkasInfo = bResult.ok
      ? " | Berkas: ✅ " + bResult.mode
      : " | Berkas: ⚠️ " + bResult.info;
  } catch(bErr) {
    berkasInfo = " | Berkas: ⚠️ " + bErr.message;
  }

  // Generate KTM
  let ktmInfo = "";
  try {
    const ktmUrl = generateKTM(nim, data, tahun);
    ktmInfo = " | KTM: " + ktmUrl;
  } catch(ktmErr) {
    ktmInfo = " | ⚠️ KTM gagal: " + ktmErr.message;
  }

  sheet.getRange(row, C.CATATAN + 1)
       .setValue("✅ NIM: " + nim + " | DB: OK" + berkasInfo + ktmInfo +
                 " | " + new Date().toLocaleString("id-ID"));
}

function getNoUrutProdi(prodi, tahun) {
  const prefix = tahun + String(KODE_PRODI_NIM[prodi]);
  try {
    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName("mahasiswa");
    if (!sheet) return 1;
    const data  = sheet.getDataRange().getValues();
    let count = 0;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).startsWith(prefix)) count++;
    }
    return count + 1;
  } catch(e) { return 1; }
}

// ═══════════════════════════════════════════════════════════════
// DATABASE MAHASISWA
// ═══════════════════════════════════════════════════════════════

function inputKeMahasiswaDB(nim, data, tahun) {
  const db       = SpreadsheetApp.openById(DB_SHEET_ID);
  const ts       = new Date();
  const prodi    = String(data[C.PRODI]).trim().toUpperCase();
  const idDaftar = "PMB-20"+tahun+"-"+String(data[C.NO_URUT]).padStart(4,"0");

  const shM = db.getSheetByName("mahasiswa");
  shM.appendRow([
    nim, data[C.NIK], idDaftar, data[C.NAMA], data[C.JK], data[C.TEMPAT], data[C.TGL_LHR],
    PRODI_DB_ID[prodi] || "", FAKULTAS_DB_ID[prodi] || "",
    data[C.JENJANG] || "S1", "20"+tahun, "20"+tahun, "Aktif",
    data[C.EMAIL], data[C.F34] || data[C.F23], ts, ts
  ]);

  const shK = db.getSheetByName("kontak_mahasiswa");
  shK.appendRow([shK.getLastRow(), nim, data[C.NO_HP], data[C.EMAIL], data[C.ALAMAT]]);

  const shF = db.getSheetByName("keluarga_mahasiswa");
  shF.appendRow([shF.getLastRow(), nim,
    data[C.AYAH], data[C.PKJ_AYAH], data[C.IBU], data[C.PKJ_IBU], data[C.ALAMAT_ORTU]]);

  const shA = db.getSheetByName("akademik_mahasiswa");
  shA.appendRow([shA.getLastRow(), nim, 0, 0, 1, "Aktif"]);

  const shB = db.getSheetByName("berkas_mahasiswa");
  shB.appendRow([shB.getLastRow(), nim,
    data[C.IJAZAH], data[C.KTP], data[C.KK],
    data[C.F23], data[C.F34], data[C.F46],
    data[C.KHS] || "", data[C.SURAT] || "", ts]);
}

// ═══════════════════════════════════════════════════════════════
// GENERATE KTM
// ═══════════════════════════════════════════════════════════════

function generateKTM(nim, data, tahun) {
  const prodi      = String(data[C.PRODI]).trim().toUpperCase();
  const jenjang    = String(data[C.JENJANG]).trim() || "S1";
  const nama       = String(data[C.NAMA]).trim();
  const tahunMasuk = "20" + tahun;
  const tahunAkhir = String(parseInt(tahunMasuk) + (jenjang === "S2" ? 3 : 4));

  // Resolusi fakultas
  const fakultasId  = FAKULTAS_DB_ID[prodi] || 1;
  const namaFakultas = NAMA_FAKULTAS[fakultasId] || "Fakultas " + fakultasId;

  // Format tempat & tanggal lahir → "Jakarta, 15 Januari 2000"
  const tempatLahir = String(data[C.TEMPAT]).trim();
  const tglRaw      = data[C.TGL_LHR];
  const tglFormatted = formatTanggalIndonesia(tglRaw);
  const ttl         = tempatLahir + ", " + tglFormatted;

  // ── Salin template ke Drive ──────────────────────────────────
  const folderKTM  = DriveApp.getFolderById(KTM_FOLDER_ID);
  const folderThn  = getOrCreateFolder(folderKTM, tahunMasuk);
  const folderProd = getOrCreateFolder(folderThn, prodi);
  const namaFile   = "KTM_" + nim + "_" + nama.replace(/\s+/g,"_");

  const ktmFile = DriveApp.getFileById(KTM_TEMPLATE_ID).makeCopy(namaFile, folderProd);
  ktmFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // ── Buka Presentasi ──────────────────────────────────────────
  const pres  = SlidesApp.openById(ktmFile.getId());
  const slide = pres.getSlides()[0];

  // ── Replace semua placeholder teks ──────────────────────────
  const replacements = {
    "{{NIM}}":          nim,
    "{{NAMA}}":         nama,
    "{{PRODI}}":        NAMA_PRODI_LENGKAP[prodi] || prodi,
    "{{FAKULTAS}}":     namaFakultas,
    "{{JENJANG}}":      jenjang,
    "{{TTL}}":          ttl,
    "{{ANGKATAN}}":     tahunMasuk,
    "{{MASA_BERLAKU}}": tahunMasuk + " s.d. " + tahunAkhir
  };
  for (const [ph, val] of Object.entries(replacements)) {
    pres.replaceAllText(ph, val);
  }



  // ── Insert foto mahasiswa ────────────────────────────────────
  // Foto dibungkus try-catch tersendiri agar kegagalan foto
  // tidak membatalkan teks yang sudah tersimpan.
  try {
    const foto = getPhotoBlob(data);
    if (foto.blob) {
      insertFotoKTM(slide, foto.blob, nim);
    }
  } catch(fotoErr) {
    // Foto gagal tidak apa-apa — KTM tetap valid dengan teks lengkap
    Logger.log("Info: Foto KTM gagal diinsert untuk NIM " + nim + ": " + fotoErr.message);
  }

  return ktmFile.getUrl();
}

// ── FORMAT TANGGAL KE BAHASA INDONESIA ──────────────────────────
// Input: Date object, string "YYYY-MM-DD", atau nilai mentah dari Sheets
// Output: "15 Januari 2000"
function formatTanggalIndonesia(tglRaw) {
  const BULAN = [
    "Januari","Februari","Maret","April","Mei","Juni",
    "Juli","Agustus","September","Oktober","November","Desember"
  ];

  let d;
  if (tglRaw instanceof Date) {
    d = tglRaw;
  } else {
    // Sheets kadang menyimpan tanggal sebagai string "YYYY-MM-DD" atau "MM/DD/YYYY"
    d = new Date(tglRaw);
  }

  // Validasi — jika tidak bisa di-parse, kembalikan string asli
  if (isNaN(d.getTime())) return String(tglRaw);

  return d.getDate() + " " + BULAN[d.getMonth()] + " " + d.getFullYear();
}

// ── INSERT FOTO KE SLIDE ─────────────────────────────────────────
// Mencari placeholder image dengan Alt Text (Title atau Description) = "FOTO_MHS"
// Jika tidak ditemukan, sisipkan di posisi default kiri atas
function insertFotoKTM(slide, blob, nim) {
  const elements = slide.getPageElements();

  for (const el of elements) {
    if (el.getPageElementType() !== SlidesApp.PageElementType.IMAGE) continue;
    const img = el.asImage();

    // FIX: cek KEDUANYA — Title dan Description dari Alt Text
    const altTitle = img.getTitle()       || "";
    const altDesc  = img.getDescription() || "";

    if (altTitle !== "FOTO_MHS" && altDesc !== "FOTO_MHS") continue;

    // Simpan posisi & ukuran placeholder, lalu hapus & ganti
    const left = el.getLeft(), top = el.getTop();
    const w    = el.getWidth(), h   = el.getHeight();
    el.remove();

    const newImg = slide.insertImage(blob, left, top, w, h);
    newImg.setTitle("foto_ktm");
    newImg.setDescription(nim);
    return; // selesai — hanya ganti satu placeholder
  }

  // Fallback: placeholder tidak ditemukan → sisipkan di posisi default
  // Ukuran 2.5cm x 3.35cm dikonversi ke points (1 cm = 28.346 pt)
  const newImg = slide.insertImage(blob, 14, 60, 71, 95);
  newImg.setTitle("foto_ktm");
  newImg.setDescription(nim);
}

// ── PEMILIHAN FOTO ────────────────────────────────────────────────
// Prioritas: foto_3x4 → foto_2x3 → foto_4x6
// Antisipasi foto sama: URL identik → cukup dicoba sekali (Set dedup)
// Verifikasi MIME type → pastikan gambar, bukan PDF
function getPhotoBlob(data) {
  const candidates = [
    { url: String(data[C.F34]).trim(), label:"3x4" },
    { url: String(data[C.F23]).trim(), label:"2x3" },
    { url: String(data[C.F46]).trim(), label:"4x6" }
  ];

  const tried = new Set();
  for (const c of candidates) {
    if (!c.url || c.url === "undefined" || tried.has(c.url)) continue;
    tried.add(c.url);

    const fileId = extractDriveFileId(c.url);
    if (!fileId) continue;

    try {
      const file = DriveApp.getFileById(fileId);
      if (!file.getMimeType().startsWith("image/")) continue;
      return { blob: file.getBlob(), source: c.label };
    } catch(_) { continue; }
  }
  return { blob:null, source:null };
}

function extractDriveFileId(url) {
  if (!url) return null;
  const patterns = [
    /\/file\/d\/([a-zA-Z0-9_-]{10,})/,
    /[?&]id=([a-zA-Z0-9_-]{10,})/,
    /\/d\/([a-zA-Z0-9_-]{10,})\//
  ];
  for (const p of patterns) {
    const m = url.match(p);
    if (m) return m[1];
  }
  return null;
}

// ── DEBUGGER — jalankan manual dari Apps Script Editor ──────────
// Menampilkan semua Alt Text elemen gambar di slide 0 template.
// Gunakan ini untuk memastikan placeholder "FOTO_MHS" sudah benar.
function debugTemplateAltText() {
  const pres    = SlidesApp.openById(KTM_TEMPLATE_ID);
  const slide   = pres.getSlides()[0];
  const elements = slide.getPageElements();
  const hasil = [];

  for (const el of elements) {
    const type = el.getPageElementType();
    if (type === SlidesApp.PageElementType.IMAGE) {
      const img = el.asImage();
      hasil.push("IMAGE | Title: \"" + img.getTitle() +
                 "\" | Description: \"" + img.getDescription() + "\"");
    } else if (type === SlidesApp.PageElementType.SHAPE) {
      const shape = el.asShape();
      const txt   = shape.getText().asRenderedString().substring(0,40);
      hasil.push("SHAPE | text: \"" + txt + "\"");
    }
  }

  Logger.log("=== ELEMEN DI SLIDE 0 TEMPLATE ===");
  hasil.forEach(h => Logger.log(h));
  Logger.log("=== TOTAL: " + hasil.length + " elemen ===");
}

// ═══════════════════════════════════════════════════════════════
// UTILITAS — Proses ulang data lama yang sudah Terverifikasi
// ═══════════════════════════════════════════════════════════════

function prosesUlangSemuaTerverifikasi() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("master_pmb");
  const data  = sheet.getDataRange().getValues();
  let count   = 0;
  for (let i = 1; i < data.length; i++) {
    const status = String(data[i][C.STATUS]).trim();
    const nim    = String(data[i][C.NIM]).trim();
    if (status === "Terverifikasi" && !nim) {
      prosesVerifikasi(sheet, i + 1);
      count++;
      Utilities.sleep(800);
    }
  }
  Logger.log("Selesai. Total: " + count + " mahasiswa diproses.");
}

// ═══════════════════════════════════════════════════════════════
// DEBUG FOTO — Jalankan dari Apps Script Editor untuk diagnosa
// Ganti angka 2 dengan nomor baris mahasiswa di master_pmb
// ═══════════════════════════════════════════════════════════════

function debugFotoMahasiswa(targetRow) {
  const row   = targetRow || 2; // default baris 2 (mahasiswa pertama)
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("master_pmb");
  const data  = sheet.getRange(row, 1, 1, C.NIM + 1).getValues()[0];

  const namaCell = String(data[C.NAMA]);
  Logger.log("=== DEBUG FOTO — Baris " + row + " | " + namaCell + " ===");

  const slots = [
    { label:"F34 (3x4)", url: String(data[C.F34]).trim() },
    { label:"F23 (2x3)", url: String(data[C.F23]).trim() },
    { label:"F46 (4x6)", url: String(data[C.F46]).trim() },
  ];

  for (const s of slots) {
    Logger.log("\n--- " + s.label + " ---");
    Logger.log("URL: " + (s.url || "(kosong)"));

    if (!s.url) { Logger.log("→ SKIP: URL kosong"); continue; }

    const fileId = extractDriveFileId(s.url);
    Logger.log("File ID: " + (fileId || "(tidak terparsing!)"));
    if (!fileId) continue;

    try {
      const file     = DriveApp.getFileById(fileId);
      const mime     = file.getMimeType();
      const size     = file.getSize();
      const isImg    = mime.startsWith("image/");
      Logger.log("Nama file: " + file.getName());
      Logger.log("MIME type: " + mime + (isImg ? " ✅" : " ❌ BUKAN GAMBAR!"));
      Logger.log("Ukuran: " + (size / 1024).toFixed(1) + " KB");
      Logger.log("Owner: " + file.getOwner().getEmail());

      if (isImg) {
        const blob = file.getBlob();
        Logger.log("Blob berhasil diambil ✅ — " + s.label + " akan dipakai untuk KTM");
        break;
      }
    } catch(e) {
      Logger.log("→ ERROR getFileById: " + e.message);
    }
  }
  Logger.log("\n=== SELESAI ===");
}

// ═══════════════════════════════════════════════════════════════
// MANAJEMEN BERKAS MAHASISWA TERVERIFIKASI
//
// Saran arsitektur:
//   Opsi A — Shortcut (RECOMMENDED):
//     File tetap di PMB, dibuat shortcut di Database Mahasiswa.
//     Aman: tidak ada risiko file hilang di tengah proses.
//     URL berkas di spreadsheet tetap valid.
//
//   Opsi B — Pindah fisik (Move):
//     Folder personal mahasiswa dipindah dari PMB ke Database Mahasiswa.
//     File ID tidak berubah → URL di spreadsheet tetap valid.
//     Cocok jika ingin PMB folder bersih setelah verifikasi.
//
// Struktur setelah diproses:
//   Database Mahasiswa/
//     Berkas Mahasiswa Terverifikasi/
//       2025/
//         PAI/
//           252100001_AHMAD_FAUZI/   ← folder personal (dipindah atau shortcut)
//             Ijazah_SMA_...
//             KTP_...
//             KK_...
//             Foto_2x3_...
//             Foto_3x4_...
//             Foto_4x6_...
// ═══════════════════════════════════════════════════════════════

// Dipanggil dari prosesVerifikasi setelah NIM di-generate
function kelolaFolderBerkas(nim, data, tahun) {
  const prodi = String(data[C.PRODI]).trim().toUpperCase();
  const nama  = String(data[C.NAMA]).trim();
  const nisn  = String(data[C.NISN]).trim();

  // Nama folder personal mahasiswa (sesuai pola di uploadSingleFile)
  const namaPersonal = nama.replace(/\s+/g,"_") + "_" + nisn;

  // Tentukan subfolder sumber di PMB
  const statusMhs    = String(data[C.STAT_MHS]).trim();
  const subNamaPMB   = (statusMhs === "Pindahan") ? "Mahasiswa Pindahan" : "Mahasiswa Baru";

  try {
    // Cari folder personal di PMB
    const folderPMBRoot = DriveApp.getFolderById(FOLDER_ID);
    const folderSub     = cariSubFolder(folderPMBRoot, subNamaPMB);
    if (!folderSub) {
      Logger.log("Folder PMB sub tidak ditemukan: " + subNamaPMB);
      return { ok:false, info:"Folder PMB '" + subNamaPMB + "' tidak ditemukan" };
    }

    const folderPersonal = cariSubFolder(folderSub, namaPersonal);
    if (!folderPersonal) {
      Logger.log("Folder personal tidak ditemukan: " + namaPersonal);
      return { ok:false, info:"Folder personal '" + namaPersonal + "' tidak ditemukan" };
    }

    // Buat struktur tujuan di Database Mahasiswa
    // Berkas Mahasiswa Terverifikasi / [tahun] / [prodi] / [nim]_[nama]
    const folderBerkasDB  = DriveApp.getFolderById(BERKAS_DB_FOLDER_ID);
    const folderTahun     = getOrCreateFolder(folderBerkasDB, "20" + tahun);
    const folderProdi     = getOrCreateFolder(folderTahun, prodi);
    const namaTujuan      = nim + "_" + nama.replace(/\s+/g,"_");

    // ── Opsi A: Buat Shortcut (RECOMMENDED) ────────────────────
    // Shortcut menunjuk ke folder personal yang ada di PMB.
    // Folder PMB tidak berubah, hanya ada "alias" di Database.
    const shortcut = folderProdi.createShortcut(folderPersonal.getId());
    shortcut.setName(namaTujuan);
    Logger.log("Shortcut berkas dibuat: " + shortcut.getUrl());
    return { ok:true, info:"Shortcut: " + shortcut.getUrl(), mode:"shortcut" };

    // ── Opsi B: Pindah fisik (uncomment jika ingin move) ───────
    // PERINGATAN: Tidak bisa di-undo otomatis jika ada error setelahnya.
    // folderPersonal.setName(namaTujuan);  // rename sekalian jadi NIM_NAMA
    // folderPersonal.moveTo(folderProdi);
    // return { ok:true, info:"Dipindah ke: " + folderPersonal.getUrl(), mode:"move" };

  } catch(e) {
    return { ok:false, info:"Error: " + e.message };
  }
}

// Helper: cari subfolder by name (tidak throw jika tidak ada)
function cariSubFolder(parentFolder, namaTarget) {
  const it = parentFolder.getFoldersByName(namaTarget);
  return it.hasNext() ? it.next() : null;
}

// kelolaFolderBerkas sudah diintegrasikan ke prosesVerifikasi() di atas.