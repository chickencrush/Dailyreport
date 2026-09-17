# Daily Report API PWA

## 1. Update Apps Script (wajib dahulu)
- Buka project Daily Report yang lama.
- Ganti seluruh isi file Script Code.gs/Kode.gs dengan Code.gs paket ini. Pastikan hanya satu file backend berisi fungsi yang sama.
- Ganti seluruh isi HTML Index dengan Index.html huruf besar paket ini. File ini mempertahankan akses langsung /exec.
- Konfigurasi spreadsheet dan email admin sudah terisi. Database lama tidak perlu setupApp ulang.
- Simpan. Deploy > Manage deployments > Edit > New version.
- Execute as: Me. Who has access: Anyone (bukan hanya pengguna yang login akun Google). API memakai login internal dan token aplikasi untuk membatasi data. Bila Anyone tidak tersedia karena kebijakan organisasi, API langsung dari GitHub ini memerlukan hosting/proxy lain.
- Deploy. Pertahankan deployment yang sama agar URL lama tidak berubah. Jika URL berubah, update config.js di GitHub.

## 2. Update GitHub
Unggah file hasil ekstrak langsung ke akar repository Dailyreport. index.html HURUF KECIL adalah aplikasi GitHub; Index.html HURUF BESAR untuk Apps Script.
File web yang wajib: index.html, config.js, sw.js, manifest.webmanifest, icon-180.png, icon-192.png, icon-512.png. Paket juga menyertakan backend dan tes.
Settings > Pages > Deploy from a branch > main > /(root) > Save.
Halaman aplikasi: https://chickencrush.github.io/Dailyreport/
Jika halaman pembungkus lama masih terlihat, Chrome PC: Ctrl+Shift+R. Jika masih tertahan service worker lama, DevTools > Application > Service Workers > Unregister, lalu reload. Versi service worker baru menghapus cache Daily Report lama setelah aktif.

## 3. Pasang
PC Chrome/Edge atau Android Chrome: tombol Pasang Daily Report atau menu browser > Pasang aplikasi. iOS Safari: Bagikan > Tambahkan ke Layar Utama. Ikon menggunakan logo yang diberikan.

## Cara kerja
Frontend dimuat dari GitHub Pages tanpa iframe. Kedua file HTML mendeteksi konteks: ketika berjalan di /exec, koneksi memakai google.script.run; di GitHub, koneksi memakai API POST. Backend mengembalikan JSON dari doPost; frontend mengirim POST text/plain dan mengikuti redirect Google. Token/password tidak dikirim di URL. Data pengguna tetap memakai otorisasi backend. Operasi backend dibatasi daftar API yang eksplisit; setupApp dan helper privat tidak dapat dipanggil melalui API.
Status laporan baru Hold > Proses > Selesai. Link file foto Google Drive wajib saat selesai; format diperiksa, izin akses/isi foto tidak diverifikasi. Bukti disimpan terpisah dari lampiran.
Backend tidak dicache service worker. Data laporan, login, upload dan export memerlukan internet. Shell aplikasi dapat dimuat offline.

## Pengujian
npm test mencakup sintaks, simulasi frontend, daftar/login backend, status, bukti, hak akses, whitelist API, token POST dan penanganan error. Browser lintas origin, redirect/CORS pada deployment Google nyata dan instalasi perangkat fisik belum diuji. Jika koneksi API gagal, cek akses Anyone dan versi deployment terlebih dahulu. Jangan gunakan mode no-cors, karena respons tidak dapat dibaca.
Referensi API: https://developers.google.com/apps-script/guides/content
