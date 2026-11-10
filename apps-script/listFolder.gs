/**
 * Tambahan untuk Google Apps Script galeri: mengembalikan daftar foto di sebuah
 * folder Google Drive, supaya halaman bisa menampilkan grid foto sendiri
 * (lebih rapi daripada tampilan bawaan Drive).
 *
 * CARA PASANG
 * 1. Buka proyek Apps Script yang URL-nya ada di GOOGLE_APPS_SCRIPT_URL.
 * 2. Tempel fungsi listFolder_ di bawah ini ke proyek tersebut.
 * 3. Di AWAL fungsi doGet(e) yang sudah ada, tambahkan satu baris:
 *
 *      if (e && e.parameter && e.parameter.action === 'listFolder') return listFolder_(e.parameter.id);
 *
 * 4. Simpan, lalu Deploy > Manage deployments > (ikon pensil) > Version: New version > Deploy.
 *    (Memakai "Manage deployments" membuat URL /exec tetap sama.)
 *    Saat diminta, izinkan akses Google Drive.
 *
 * Foto di folder tetap harus dibagikan "Siapa saja yang memiliki link" agar bisa
 * tampil di halaman.
 */
function listFolder_(folderId) {
  var out;
  try {
    if (!folderId || !/^[a-zA-Z0-9_-]+$/.test(folderId)) throw new Error('ID folder tidak valid');
    var it = DriveApp.getFolderById(folderId).getFiles();
    var files = [];
    while (it.hasNext()) {
      var f = it.next();
      var mime = f.getMimeType();
      if (mime.indexOf('image/') !== 0) continue;
      files.push({
        id: f.getId(),
        name: f.getName(),
        mimeType: mime,
        created: f.getDateCreated().toISOString()
      });
    }
    // Urutkan dari yang paling lama diunggah
    files.sort(function (a, b) { return a.created < b.created ? -1 : a.created > b.created ? 1 : 0; });
    out = { status: 'success', files: files };
  } catch (err) {
    out = { status: 'error', message: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
