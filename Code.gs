const SHEET_ID  = "1Cq6YKHMx7RyrNXhwFzV76Eja3QW8hDJIk0Ss1AYjzHU";
const FOLDER_ID = "1qC1ZymS4xeuJltdbIOFmtWBFlZXy3TZu";

function doGet(e) {
  const tpl = HtmlService.createTemplateFromFile('Form');
  tpl.cacheBust = new Date().getTime();
  // Ambil kode referral dari URL parameter ?ref=KODE (jika ada)
  tpl.refCode = (e && e.parameter && e.parameter.ref)
    ? e.parameter.ref.toString().toUpperCase().replace(/[^A-Z0-9\-_]/g, '').substring(0, 20)
    : '';

  return tpl.evaluate()
    .setTitle('PMB 2025/2026 — IAI Al-Aqidah Al-Hasyimiyyah Jakarta')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Cek duplikat NIK / NISN di sheet Master Data
// Dipanggil dari form (cekDuplikatForm) dan dari simpanData sebagai safety net
function cekDuplikat(nik, nisn) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('master_pmb');
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {           // i=0 adalah header
      const rowNik  = String(data[i][2]).trim();       // Kolom C = NIK
      const rowNisn = String(data[i][7]).trim();       // Kolom H = NISN
      const rowNama = String(data[i][3]).trim();       // Kolom D = Nama
      if (rowNik === nik || rowNisn === nisn) {
        return { duplikat: true, baris: i + 1, nama: rowNama };
      }
    }
    return { duplikat: false };
  } catch (e) {
    // Jika gagal cek, biarkan lanjut — akan ditangkap simpanData
    return { duplikat: false };
  }
}

// Wrapper yang dipanggil langsung dari form sebelum pindah step
function cekDuplikatForm(nik, nisn) {
  return cekDuplikat(nik.trim(), nisn.trim());
}

// Upload SATU file, kembalikan URL Drive
function uploadSingleFile(b64, fileName, mimeType, subFolderName, personalFolderName) {
  try {
    const root   = DriveApp.getFolderById(FOLDER_ID);
    const sub    = getOrCreateFolder(root, subFolderName);
    const person = getOrCreateFolder(sub, personalFolderName);
    const blob   = Utilities.newBlob(Utilities.base64Decode(b64), mimeType, fileName);
    const file   = person.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { success: true, url: file.getUrl() };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// Simpan data ke Sheets (dipanggil terakhir setelah semua file terupload)
function simpanData(fd) {
  try {
    // ── Safety net: cek duplikat sekali lagi (mencegah race condition double-submit) ──
    const cek = cekDuplikat(fd.nik.trim(), fd.nisn.trim());
    if (cek.duplikat) {
      return {
        success: false, duplikat: true,
        message: 'NIK/NISN sudah terdaftar (baris ' + cek.baris + ', ' + cek.nama + '). Hubungi panitia.'
      };
    }

    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('master_pmb');
    const no    = sheet.getLastRow(); // nomor urut
    const ts    = new Date();

    sheet.appendRow([
      no,
      ts,
      fd.nik,
      fd.namaLengkap,
      fd.tempatLahir,
      fd.tanggalLahir,
      fd.jenisKelamin,
      fd.nisn,
      fd.sekolahAsal,
      fd.prodiPilihan,
      fd.jenjang       || 'S1',  // Jenjang (S1 default, S2 hanya untuk PAI)
      fd.kelas,
      fd.statusMahasiswa,
      fd.ptAsal       || '',
      fd.jenjangAsal  || '',
      fd.prodiAsal    || '',
      fd.alamatDomisili,
      fd.email,
      fd.noHp          || '',   // No. HP / WhatsApp
      fd.namaBapak,
      fd.pekerjaanBapak,
      fd.namaIbu,
      fd.pekerjaanIbu,
      fd.alamatOrtu,
      fd.urlIjazah,
      fd.urlKTP,
      fd.urlKK,
      fd.urlFoto23,
      fd.urlFoto34,
      fd.urlFoto46,
      fd.urlKHS        || '',
      fd.urlSurat      || '',
      fd.kodeReferral  || '',   // Kode Referral
      'Pending',
      ''
    ]);

    return { success: true, noPendaftaran: 'PMB-2025-' + String(no).padStart(4, '0') };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getOrCreateFolder(parent, name) {
  const existing = parent.getFoldersByName(name);
  return existing.hasNext() ? existing.next() : parent.createFolder(name);
}

// ── REFERRAL ────────────────────────────────────────────────────────────────

// Validasi kode referral saat diketik manual di form
function cekReferral(kode) {
  if (!kode) return { valid: false, nama: '' };
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('referral');
    if (!sheet) return { valid: false, nama: '' };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {          // baris 1 = header
      if (String(data[i][0]).trim().toUpperCase() === kode.toUpperCase() && data[i][2] === true) {
        return { valid: true, nama: data[i][1] };     // kolom A=Kode, B=Nama Agen, C=Aktif
      }
    }
    return { valid: false, nama: '' };
  } catch (e) {
    return { valid: false, nama: '' };
  }
}

// Hitung total referral per kode (dipanggil dari dashboard admin jika perlu)
function rekapReferral() {
  const masterSheet   = SpreadsheetApp.openById(SHEET_ID).getSheetByName('master_pmb');
  const referralSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('referral');
  if (!masterSheet || !referralSheet) return;

  const masterData  = masterSheet.getDataRange().getValues();
  // Kolom Kode Referral = indeks 32 (kolom AG, 0-based) — sesuaikan jika beda
  const KOLOM_REF    = 32;
  const KOLOM_STATUS = 33; // kolom Status (Pending/Verified/dll)

  const rekap = {};
  for (let i = 1; i < masterData.length; i++) {
    const kode   = String(masterData[i][KOLOM_REF]).trim().toUpperCase();
    const status = String(masterData[i][KOLOM_STATUS]).trim();
    if (!kode) continue;
    if (!rekap[kode]) rekap[kode] = { total: 0, verified: 0 };
    rekap[kode].total++;
    if (status === 'Verified') rekap[kode].verified++;
  }

  // Tulis rekap ke sheet Referral kolom D & E
  const refData = referralSheet.getDataRange().getValues();
  for (let i = 1; i < refData.length; i++) {
    const kode = String(refData[i][0]).trim().toUpperCase();
    const r    = rekap[kode] || { total: 0, verified: 0 };
    referralSheet.getRange(i + 1, 4).setValue(r.total);      // D = Total Daftar
    referralSheet.getRange(i + 1, 5).setValue(r.verified);   // E = Sudah Verified
  }
}