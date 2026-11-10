/**
 * Apps Script untuk Galeri Romantis & Google Sheets Data Sync
 *
 * PETUNJUK PENYETELAN (pertama kali):
 * 1. Buka Google Sheets baru di Google Drive Anda.
 * 2. Klik menu Ekstensi -> Apps Script.
 * 3. Hapus semua kode default, lalu tempelkan seluruh kode ini.
 * 4. Klik "Deploy" (Terapkan) -> "Deployment baru" (New deployment).
 * 5. Pilih jenis (Select type): "Aplikasi Web" (Web App).
 * 6. Setel "Jalankan sebagai" (Execute as): "Saya" (Me).
 * 7. Setel "Siapa yang memiliki akses" (Who has access): "Siapa saja" (Anyone).
 * 8. Klik "Terapkan" (Deploy) & berikan izin otorisasi (Authorize Access).
 * 9. Salin URL Web App (format: https://script.google.com/macros/s/.../exec).
 * 10. Tempelkan URL tersebut ke variabel GOOGLE_APPS_SCRIPT_URL di index.html.
 *
 * MEMPERBARUI KODE (sudah pernah di-deploy):
 * - Ganti seluruh isi Code.gs dengan kode ini, simpan.
 * - Deploy -> Kelola deployment (Manage deployments) -> ikon pensil ->
 *   Versi: "Versi baru" (New version) -> Deploy.
 *   Dengan cara ini URL /exec TIDAK berubah, jadi index.html tidak perlu diedit.
 * - Izinkan akses Google Drive saat diminta (dibutuhkan untuk daftar foto
 *   folder dan untuk menyimpan foto yang diunggah dari HP/PC).
 *
 * FITUR
 * - GET  ?               -> semua data (albums, photos, anniversaryDate)
 * - GET  ?action=listFolder&id=FOLDER_ID -> daftar foto di folder Drive
 * - POST {action:'syncAll', data:{...}}  -> simpan semua data
 *   Foto yang diunggah dari HP/PC (base64) otomatis disimpan sebagai file di
 *   Google Drive (folder UPLOAD_FOLDER_NAME), karena satu sel Sheets hanya
 *   muat 50.000 karakter. Yang disimpan di Sheets cukup link-nya.
 */

const UPLOAD_FOLDER_NAME = 'Galeri Kenangan - Upload';
const FOLDER_CACHE_SECONDS = 300; // daftar foto folder disimpan 5 menit agar cepat

const ALBUM_HEADERS = ['id', 'title', 'date', 'location', 'driveUrl', 'driveFolderId', 'coverUrl', 'description'];
const PHOTO_HEADERS = ['id', 'title', 'url', 'date', 'category', 'location', 'caption', 'isFavorite'];
const SETTING_HEADERS = ['key', 'value'];

// ======================================================================
// GET
// ======================================================================
function doGet(e) {
  try {
    const action = e && e.parameter && e.parameter.action;
    if (action === 'listFolder') return createJsonResponse(listFolder(e.parameter.id));

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const albumsSheet = getOrCreateSheet(ss, 'Albums', ALBUM_HEADERS);
    const photosSheet = getOrCreateSheet(ss, 'Photos', PHOTO_HEADERS);
    const settingsSheet = getOrCreateSheet(ss, 'Settings', SETTING_HEADERS);

    const albums = readSheetObjects(albumsSheet);
    const photos = readSheetObjects(photosSheet).map(function (p) {
      p.isFavorite = p.isFavorite === true || String(p.isFavorite).toLowerCase() === 'true';
      return p;
    });

    let anniversaryDate = null;
    const annivSetting = readSheetObjects(settingsSheet).find(function (s) { return s.key === 'anniversaryDate'; });
    if (annivSetting) anniversaryDate = annivSetting.value;

    return createJsonResponse({
      status: 'success',
      data: { albums: albums, photos: photos, anniversaryDate: anniversaryDate }
    });
  } catch (err) {
    return createJsonResponse({ status: 'error', message: err.toString() });
  }
}

// ======================================================================
// POST
// ======================================================================
function doPost(e) {
  // Cegah dua penyimpanan berjalan bersamaan dan saling menimpa
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return createJsonResponse({ status: 'error', message: 'Server sibuk, coba lagi.' });
  }

  try {
    if (!e || !e.postData || !e.postData.contents) {
      return createJsonResponse({ status: 'error', message: 'No payload received' });
    }

    const contents = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (contents.action === 'syncAll' && contents.data) {
      const data = contents.data;
      const urlMap = {};

      if (Array.isArray(data.albums)) {
        const albumsSheet = getOrCreateSheet(ss, 'Albums', ALBUM_HEADERS);
        writeSheetObjects(albumsSheet, ALBUM_HEADERS, data.albums);
      }

      if (Array.isArray(data.photos)) {
        const photos = data.photos.map(function (p) {
          const copy = Object.assign({}, p);
          if (typeof copy.url === 'string' && copy.url.indexOf('data:image/') === 0) {
            const newUrl = saveBase64ToDrive(copy.id, copy.url);
            if (newUrl) {
              copy.url = newUrl;
              urlMap[copy.id] = newUrl;
            } else {
              copy.url = ''; // jangan tulis base64 raksasa ke sel
            }
          }
          return copy;
        });
        const photosSheet = getOrCreateSheet(ss, 'Photos', PHOTO_HEADERS);
        writeSheetObjects(photosSheet, PHOTO_HEADERS, photos);
      }

      if (data.anniversaryDate) {
        const settingsSheet = getOrCreateSheet(ss, 'Settings', SETTING_HEADERS);
        writeSheetObjects(settingsSheet, SETTING_HEADERS, [{ key: 'anniversaryDate', value: data.anniversaryDate }]);
      }

      return createJsonResponse({
        status: 'success',
        message: 'Data synced successfully to Google Sheets',
        urlMap: urlMap // id foto -> link Drive baru, supaya halaman berhenti mengirim base64
      });
    }

    return createJsonResponse({ status: 'error', message: 'Invalid action payload' });
  } catch (err) {
    return createJsonResponse({ status: 'error', message: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

// ======================================================================
// Folder Drive: daftar foto
// ======================================================================
function listFolder(folderId) {
  if (!folderId || !/^[a-zA-Z0-9_-]+$/.test(folderId)) {
    return { status: 'error', message: 'ID folder tidak valid' };
  }

  const cache = CacheService.getScriptCache();
  const cacheKey = 'folder_' + folderId;
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  try {
    const it = DriveApp.getFolderById(folderId).getFiles();
    const files = [];
    while (it.hasNext()) {
      const f = it.next();
      const mime = f.getMimeType();
      if (mime.indexOf('image/') !== 0) continue;
      files.push({
        id: f.getId(),
        name: f.getName(),
        mimeType: mime,
        created: f.getDateCreated().toISOString()
      });
    }
    // Urut dari yang paling lama diunggah
    files.sort(function (a, b) { return a.created < b.created ? -1 : a.created > b.created ? 1 : 0; });

    const result = { status: 'success', files: files };
    const json = JSON.stringify(result);
    if (json.length < 90000) cache.put(cacheKey, json, FOLDER_CACHE_SECONDS);
    return result;
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

// ======================================================================
// Simpan foto base64 (upload dari HP/PC) ke Google Drive
// ======================================================================
function saveBase64ToDrive(photoId, dataUrl) {
  try {
    const m = dataUrl.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);
    if (!m) return '';
    const mime = m[1];
    const bytes = Utilities.base64Decode(m[2]);

    // Nama file = id foto + sidik isi, jadi foto yang sama tidak diunggah dua kali
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, bytes)
      .map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('').slice(0, 10);
    const ext = mime === 'image/png' ? 'png' : mime === 'image/webp' ? 'webp' : 'jpg';
    const fileName = String(photoId).replace(/[^a-zA-Z0-9_-]/g, '') + '_' + digest + '.' + ext;

    const folder = getUploadFolder();
    const existing = folder.getFilesByName(fileName);
    const file = existing.hasNext()
      ? existing.next()
      : folder.createFile(Utilities.newBlob(bytes, mime, fileName));

    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareErr) {
      // Beberapa akun kantor/sekolah melarang berbagi publik; file tetap tersimpan
      console.warn('Gagal mengatur berbagi: ' + shareErr);
    }
    return 'https://lh3.googleusercontent.com/d/' + file.getId();
  } catch (err) {
    console.error('saveBase64ToDrive gagal: ' + err);
    return '';
  }
}

function getUploadFolder() {
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('UPLOAD_FOLDER_ID');
  if (savedId) {
    try { return DriveApp.getFolderById(savedId); } catch (e) { /* folder dihapus, buat baru */ }
  }
  const folder = DriveApp.createFolder(UPLOAD_FOLDER_NAME);
  props.setProperty('UPLOAD_FOLDER_ID', folder.getId());
  return folder;
}

// ======================================================================
// Helper Functions
// ======================================================================
function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function readSheetObjects(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(String);
  const tz = sheet.getParent().getSpreadsheetTimeZone();

  return values.slice(1)
    .filter(function (row) { return row.some(function (v) { return v !== '' && v !== null; }); })
    .map(function (row) {
      const obj = {};
      headers.forEach(function (h, i) {
        if (!h) return;
        let v = row[i];
        // Sheets bisa mengubah "2026-10-30" menjadi Date dan "1700000000000" menjadi angka.
        // Kembalikan ke teks supaya tanggal tidak bergeser dan id tetap cocok di halaman.
        if (v instanceof Date) v = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
        else if (typeof v === 'number') v = String(v);
        obj[h] = v;
      });
      return obj;
    });
}

function writeSheetObjects(sheet, headers, objects) {
  const rows = objects.map(function (obj) {
    return headers.map(function (h) {
      const v = obj[h];
      if (v === undefined || v === null) return '';
      if (typeof v === 'boolean') return v;
      return String(v);
    });
  });

  sheet.clearContents();
  const all = [headers].concat(rows);
  const range = sheet.getRange(1, 1, all.length, headers.length);
  range.setNumberFormat('@'); // simpan sebagai teks biasa
  range.setValues(all);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
}

function createJsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
