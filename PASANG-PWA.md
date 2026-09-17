# Daily Report yang dapat dipasang

1. URL Daily Report sudah terisi di docs/config.js. Tidak perlu mengisi koneksi lagi.
2. Upload isi paket ke repository GitHub. Folder docs berisi PWA; Code.gs dan Index.html di root tetap untuk Apps Script.
3. GitHub Settings > Pages > Deploy from a branch > pilih branch main dan folder /docs > Save.
4. Setelah halaman HTTPS tersedia, buka URL GitHub Pages tersebut di browser.
5. PC Chrome/Edge dan Android Chrome: Pasang. iOS Safari: Bagikan > Tambahkan ke Layar Utama. Ikon memakai logo yang diberikan.

PWA membungkus web app Apps Script dalam iframe. Backend dan data tetap menggunakan deployment lama. Bila autentikasi Google atau kebijakan browser menghalangi iframe, gunakan Buka langsung. Kegagalan akses Google di Android perlu diselesaikan pada URL /exec; PWA tidak memperbaiki izin deployment atau konflik akun Google.

Launcher tersedia offline; laporan, login dan perubahan data memerlukan internet. Ini PWA, bukan installer EXE/APK/IPA maupun publikasi App Store. Keberhasilan instalasi dan sesi dalam iframe perlu diuji setelah URL sebenarnya dan hosting tersedia.
