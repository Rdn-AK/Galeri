# Galeri
## Google Apps Script

Kode lengkap backend ada di [`apps-script/Code.gs`](apps-script/Code.gs). Petunjuk pemasangan
dan cara memperbarui deployment (tanpa mengganti URL `/exec`) ada di bagian atas file tersebut.

Fitur:
- Sinkron data album, foto, dan tanggal jadian ke Google Sheets.
- Daftar foto per folder pertemuan (`action=listFolder`) sehingga halaman bisa menampilkan
  grid foto sendiri. Tanpa ini halaman memakai tampilan bawaan Google Drive.
- Foto yang diunggah dari HP/PC disimpan sebagai file di Google Drive (folder
  "Galeri Kenangan - Upload"), karena satu sel Sheets hanya muat 50.000 karakter.
