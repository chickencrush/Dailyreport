# Chicken Crush Daily Report

## Status dan bukti
Laporan baru selalu Hold. Klik Mulai Proses untuk pindah ke Proses. Klik Selesaikan, isi link foto Google Drive, lalu Kirim Bukti & Selesaikan. Selesai adalah tahap akhir; perubahan mundur dan melompati tahap ditolak backend. Laporan lama mempertahankan status dan data. Bukti disimpan di kolom EVIDENCE_URL yang ditambahkan otomatis; lampiran DRIVE_URL tetap ada.
Link harus berupa HTTPS drive.google.com/file/d/ID atau open/uc dengan id. Sistem memeriksa format link, bukan isi foto atau izin berbagi. Pastikan foto dapat dibuka pemeriksa.

## Pasang update
Ganti SELURUH isi Code.gs dan Index.html dalam project Apps Script lama. Simpan > Deploy > Manage deployments > Edit > New version > Deploy. Tidak perlu setupApp ulang jika database sudah berjalan. Buka URL /exec.

## Upload ke GitHub
Ekstrak ZIP. Buat repository > Add file > Upload files > unggah isi paket. Git atau GitHub Desktop dapat digunakan agar folder .github ikut terunggah. Jangan commit .clasp.json atau kredensial.
GitHub menyimpan sumber; GitHub Pages tidak menjalankan backend Apps Script. Aplikasi tetap diakses dari /exec.
Opsional: gunakan clasp untuk sinkronisasi project Apps Script; .claspignore membatasi file yang dikirim ke Code.gs, Index.html dan appsscript.json.

## Pemeriksaan
npm test memeriksa simulasi frontend/backend, status berurutan, bukti wajib, penolakan akses laporan orang lain dan waktu selesai. Deployment Google dan perangkat fisik belum diuji.
