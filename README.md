1. Ini adalah sistem untuk pendaftaran mahasiswa baru
2. Sistem ini akan menyimpan data calon mahasiswa dan file lampirannya juga
3. Sistem ini menggunakan Google App Script, sehinnga semua data tersimpan pada Google Drive


FLOW
1. Calon Mahasiswa Input data pada form
2. Sistem akan menyimpan data calon mahasiswa :
    - PMB Spreadsheet untuk data pribadi
    - Folder Lampiran Mahasiswa untuk file-file lampiran
3. Setelah admin secara manual mengubah status mahasiswa menjadi Terverifikasi,
4. Sistem akan auto mengenerate KTM dan juga memasukkan data mahasiswa tersebut ke dalam database Mahasiswa (spreedsehet)


- Sistem Production : akademik@aqidah.ac.id
- Sistem Development : iai.aqidah@gmail.com



================= VAR Sistem Production (ada di Code.gs) =================
const SHEET_ID  = "1Cq6YKHMx7RyrNXhwFzV76Eja3QW8hDJIk0Ss1AYjzHU";
const FOLDER_ID = "1qC1ZymS4xeuJltdbIOFmtWBFlZXy3TZu";




================= Var sistem Development : =================
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