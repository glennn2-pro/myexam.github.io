/**
 * Download soal (CSV) untuk kelas & mapel tertentu.
 * @param {string} kelas
 * @param {string} mapel
 * @return {string} CSV string
 */
function downloadSoalCSV(kelas, mapel) {
  if (!kelas || !mapel) return 'Parameter kelas/mapel wajib diisi';
  const name = 'SOAL_' + kelas.toString().toUpperCase().trim() + '_' + mapel.toString().toUpperCase().trim();
  const sh = ss().getSheetByName(name);
  if (!sh) return 'Tidak ada sheet soal untuk kelas/mapel ini';
  const data = sh.getDataRange().getValues();
  if (!data || data.length === 0) return 'Tidak ada data';
  // Convert to CSV
  const rows = data.map(r => {
    return r.map(c => '"' + (c === undefined || c === null ? '' : ('' + c).replace(/"/g, '""')) + '"').join(',');
  });
  return rows.join('\n');
}
/**
 * Edit judul aplikasi.
 * @param {string} title - Judul baru aplikasi.
 */
function editAppTitle(title) {
  if (!title || !title.trim()) throw new Error('Judul tidak boleh kosong');
  setSetting('APP_TITLE', title.trim());
  return 'OK';
}

/**
 * Edit password admin.
 * @param {string} password - Password baru admin.
 */
function editAdminPassword(password) {
  if (!password || !password.trim()) throw new Error('Password tidak boleh kosong');
  setSetting('ADMIN_PASSWORD', password.trim());
  return 'OK';
}

/**
 * Set PIN untuk My Exams Browser (kiosk mode)
 * @param {string} pin - PIN 4-6 digit
 */
function setMyExamsBrowserPIN(pin) {
  if (!pin || !pin.trim()) throw new Error('PIN tidak boleh kosong');
  const pinStr = pin.toString().trim();
  if (!/^\d{4,6}$/.test(pinStr)) throw new Error('PIN harus 4-6 digit angka');
  setSetting('MYEXAMS_BROWSER_PIN', pinStr);
  return 'OK';
}

/**
 * Get PIN untuk My Exams Browser
 * @return {string} PIN yang tersimpan
 */
function getMyExamsBrowserPIN() {
  const pin = getSetting('MYEXAMS_BROWSER_PIN');
  return pin || '123456'; // Default PIN jika belum diset
}

/**
 * Get config untuk My Exams Browser (JSON)
 * Endpoint ini bisa diakses oleh aplikasi Android
 * @return {Object} Config object dengan PIN dan settings lain
 */
function getMyExamsBrowserConfig() {
  return {
    pin: getMyExamsBrowserPIN(),
    kiosk_mode: true,
    allow_quit: false,
    show_toolbar: false,
    app_title: getSetting('APP_TITLE') || 'Platform Ujian',
    timestamp: new Date().toISOString()
  };
}

/**
 * Get list of NPSN from LOG_LOGIN, JAWABAN_ESSAY, and HASIL sheets
 * @return {Array} Sorted array of unique NPSN values
 */
function getNPSNList() {
  try {
    const npsns = new Set();
    
    // Try to get NPSN from LOG_LOGIN sheet first
    try {
      const shLogin = ss().getSheetByName('LOG_LOGIN');
      if (shLogin) {
        const dataLogin = shLogin.getDataRange().getValues();
        if (dataLogin && dataLogin.length > 1) {
          const headersLogin = dataLogin[0].map(h => (h || '').toString().toUpperCase().replace(/\s+/g, ''));
          let npsninxLogin = headersLogin.indexOf('NPSN');
          if (npsninxLogin === -1) npsninxLogin = headersLogin.indexOf('ASAL_SEKOLAH');
          
          if (npsninxLogin !== -1) {
            for (let i = 1; i < dataLogin.length; i++) {
              const npsn = (dataLogin[i][npsninxLogin] || '').toString().trim();
              if (npsn) npsns.add(npsn);
            }
          }
        }
      }
    } catch (eLogin) {
      Logger.log('getNPSNList: error reading LOG_LOGIN: ' + eLogin);
    }
    
    // Also get NPSN from JAWABAN_ESSAY sheet as fallback/supplement
    try {
      const shEssay = ss().getSheetByName('JAWABAN_ESSAY');
      if (shEssay) {
        const dataEssay = shEssay.getDataRange().getValues();
        if (dataEssay && dataEssay.length > 1) {
          const headersEssay = dataEssay[0].map(h => (h || '').toString().toUpperCase().replace(/\s+/g, ''));
          Logger.log('getNPSNList: JAWABAN_ESSAY headers = ' + JSON.stringify(headersEssay));
          
          let npsninxEssay = headersEssay.indexOf('NPSN');
          if (npsninxEssay === -1) npsninxEssay = headersEssay.indexOf('ASAL_SEKOLAH');
          
          Logger.log('getNPSNList: NPSN index in JAWABAN_ESSAY = ' + npsninxEssay);
          
          if (npsninxEssay !== -1) {
            for (let i = 1; i < dataEssay.length; i++) {
              const npsn = (dataEssay[i][npsninxEssay] || '').toString().trim();
              if (npsn) {
                Logger.log('getNPSNList: found NPSN = ' + npsn);
                npsns.add(npsn);
              }
            }
          } else {
            Logger.log('getNPSNList: NPSN column not found in JAWABAN_ESSAY');
          }
        }
      } else {
        Logger.log('getNPSNList: JAWABAN_ESSAY sheet not found');
      }
    } catch (eEssay) {
      Logger.log('getNPSNList: error reading JAWABAN_ESSAY: ' + eEssay);
    }
    
    // Also try HASIL sheet
    try {
      const shHasil = ss().getSheetByName('HASIL');
      if (shHasil) {
        const dataHasil = shHasil.getDataRange().getValues();
        if (dataHasil && dataHasil.length > 1) {
          const headersHasil = dataHasil[0].map(h => (h || '').toString().toUpperCase().replace(/\s+/g, ''));
          let npsninxHasil = headersHasil.indexOf('NPSN');
          
          if (npsninxHasil !== -1) {
            for (let i = 1; i < dataHasil.length; i++) {
              const npsn = (dataHasil[i][npsninxHasil] || '').toString().trim();
              if (npsn) npsns.add(npsn);
            }
          }
        }
      }
    } catch (eHasil) {
      Logger.log('getNPSNList: error reading HASIL: ' + eHasil);
    }
    
    Logger.log('getNPSNList: final npsns = ' + JSON.stringify(Array.from(npsns)));
    // Return sorted array
    return Array.from(npsns).sort();
  } catch (e) {
    Logger.log('getNPSNList error: ' + e.message);
    return [];
  }
}

/**
 * Simpan jawaban essay siswa ke sheet JAWABAN_ESSAY
 * Hanya menerima jawaban ESSAY, SKIP jawaban PG
 * @param {Object} data - { nama, kelas, absen, mapel, nomor, soal, jawaban, tipe, ... }
 */
function simpanJawabanEssay(data) {
  // Validasi KETAT: hanya ESSAY yang diterima, PG dan tipe lain ditolak
  let tipeSoal = '';
  if (typeof data.tipe !== 'undefined') tipeSoal = data.tipe.toString().trim().toUpperCase();
  
  // REJECT jika tipe eksplisit adalah PG atau bukan ESSAY
  if (tipeSoal === 'PG' || (tipeSoal !== '' && tipeSoal !== 'ESSAY')) {
    Logger.log('simpanJawabanEssay: REJECTED - tipe=' + tipeSoal + ', hanya ESSAY diterima');
    return 'SKIP_PG';
  }

  if (!data || !data.nama || !data.kelas || !data.mapel || !data.nomor) throw new Error('Data essay tidak lengkap');

  // timestamp dan skor default
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var skorEssay = '';
  var skorMaksimal = 10;

  // jika tersedia info soal, coba validasi tipe dari sheet SOAL dan baca skor maksimal
  if (data.kelas && data.mapel && data.nomor) {
    try {
      const sheetName = 'SOAL_' + data.kelas.toString().toUpperCase().trim() + '_' + data.mapel.toString().toUpperCase().trim();
      const shSoal = ss().getSheetByName(sheetName);
      if (shSoal) {
        const soalData = shSoal.getDataRange().getValues();
        var foundRow = -1;
        for (let i = 1; i < soalData.length; i++) {
          if (soalData[i][0] && soalData[i][0].toString().trim() == data.nomor.toString().trim()) { foundRow = i; break; }
        }
        if (foundRow === -1) {
          Logger.log('simpanJawabanEssay: REJECTED - soal nomor ' + data.nomor + ' tidak ditemukan');
          return 'NO_SOAL';
        }
        // Temukan kolom TIPE jika ada
        let idxTipe = -1;
        const header = soalData[0].map(h => h.toString().toUpperCase().replace(/\s+/g, ''));
        for (let j = 0; j < header.length; j++) {
          if (header[j] === 'TIPE' || header[j] === 'TYPE') { idxTipe = j; break; }
        }
        const sheetTipe = (idxTipe !== -1) ? (soalData[foundRow][idxTipe] || '').toString().trim().toUpperCase() : '';
        
        // Jika sheet menunjukkan PG, REJECT
        if (sheetTipe === 'PG') {
          Logger.log('simpanJawabanEssay: REJECTED - soal nomor ' + data.nomor + ' adalah PG di sheet');
          return 'SKIP_PG';
        }
        
        // Update tipeSoal dari sheet jika belum ada
        if (!tipeSoal && sheetTipe) tipeSoal = sheetTipe;
        
        // baca skor maksimal (asumsi kolom terakhir berisi skor)
        try {
          skorMaksimal = Number(soalData[foundRow][soalData[foundRow].length - 1]) || skorMaksimal;
        } catch (e) { /* keep default */ }
      }
    } catch (e) {
      Logger.log('simpanJawabanEssay: error reading SOAL sheet: ' + e.message);
      // ignore validation errors and continue with defaults
    }
  }

  // Final check: jika masih ada indikasi PG, REJECT
  if (tipeSoal === 'PG' || (tipeSoal && tipeSoal !== 'ESSAY' && tipeSoal !== '')) {
    Logger.log('simpanJawabanEssay: FINAL REJECT - tipeSoal=' + tipeSoal);
    return 'SKIP_PG';
  }

  let sh = ss().getSheetByName('JAWABAN_ESSAY');
  if (!sh) {
    sh = ss().insertSheet('JAWABAN_ESSAY');
    sh.appendRow(['TIMESTAMP','NAMA','KELAS','ABSEN','MAPEL','NO_SOAL','SOAL','JAWABAN','SKOR','NPSN']);
  }

  if (typeof data.skor !== 'undefined') {
    skorEssay = data.skor;
  } else if (typeof data.kata_kunci !== 'undefined' && typeof data.jawaban !== 'undefined') {
    // Skor otomatis jika kata_kunci dan jawaban tersedia
    try {
      var kunciArr = (data.kata_kunci || '').toString().split(',').map(function(k){return k.trim().toLowerCase();}).filter(function(k){return k;});
      var jawabanStr = (data.jawaban || '').toString().toLowerCase();
      var matchCount = 0;
      kunciArr.forEach(function(k){ if (k && jawabanStr.indexOf(k) !== -1) matchCount++; });
      var maxSkor = (skorMaksimal !== null) ? skorMaksimal : 10;
      if (kunciArr.length > 0 && maxSkor > 0) {
        skorEssay = Math.round(matchCount / kunciArr.length * maxSkor);
      } else {
        skorEssay = '';
      }
    } catch(e) { skorEssay = ''; }
  }
  var newRow = [
    ts,
    data.nama || '',
    data.kelas || '',
    data.absen || '',
    data.mapel || '',
    data.nomor || '',
    data.soal || '',
    data.jawaban || '',
    skorEssay,
    data.npsn || ''
  ];

  // Try to find an existing row for the same student+mapel+no_soal and update it
  try {
    var all = sh.getDataRange().getValues();
    var matchRow = -1;
    var bestTs = 0;
    for (var i = 1; i < all.length; i++) {
      var row = all[i];
      var rowNama = (row[1] || '').toString().trim();
      var rowKelas = (row[2] || '').toString().trim();
      var rowMapel = (row[4] || '').toString().trim();
      var rowNo = (row[5] || '').toString().trim();
      // DEBUG: Cek apa yang di-compare
      Logger.log('simpanJawabanEssay CHECK row ' + (i+1) + ': rowNama=' + rowNama + ', rowKelas=' + rowKelas + ', rowMapel=' + rowMapel + ', rowNo=' + rowNo);
      Logger.log('simpanJawabanEssay INCOMING: nama=' + (data.nama || '').toString().trim() + ', kelas=' + (data.kelas || '').toString().trim() + ', mapel=' + (data.mapel || '').toString().trim() + ', nomor=' + (data.nomor || '').toString().trim());
      if (rowNama === (data.nama || '').toString().trim() && rowKelas === (data.kelas || '').toString().trim() && rowMapel === (data.mapel || '').toString().trim() && rowNo === (data.nomor || '').toString().trim()) {
        // choose the latest timestamped match (in case of multiple)
        var rts = 0; try { var dts = new Date(row[0]); rts = isNaN(dts.getTime()) ? 0 : dts.getTime(); } catch (e) { rts = 0; }
        if (rts >= bestTs) { bestTs = rts; matchRow = i + 1; Logger.log('simpanJawabanEssay MATCH FOUND at row ' + matchRow + ', ts=' + rts); }
      }
    }
    if (matchRow !== -1) {
      Logger.log('simpanJawabanEssay UPDATING row ' + matchRow);
      sh.getRange(matchRow, 1, 1, newRow.length).setValues([newRow]);
      return 'OK_UPDATED';
    } else {
      Logger.log('simpanJawabanEssay NO MATCH FOUND - will append new row');
    }
  } catch (e) {
    // if any error occurs, fall back to append
    Logger.log('simpanJawabanEssay error in match logic: ' + e.message);
  }

  sh.appendRow(newRow);
  Logger.log('simpanJawabanEssay APPENDED new row for nama=' + (data.nama || '') + ', nomor=' + (data.nomor || ''));
  return 'OK';
}

/**
 * Export JAWABAN_ESSAY sheet as PDF and return as base64 string.
 * Called from client via google.script.run.downloadJawabanEssayPDF()
 */
function downloadJawabanEssayPDF() {
  try {
    const sh = ss().getSheetByName('JAWABAN_ESSAY');
    if (!sh) throw new Error('Sheet JAWABAN_ESSAY tidak ditemukan');
    const gid = sh.getSheetId();
    var exportUrl = 'https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID + '/export?';
    exportUrl += 'exportFormat=pdf&format=pdf';
    exportUrl += '&gid=' + gid;
    exportUrl += '&size=A4';
    exportUrl += '&portrait=false';
    exportUrl += '&fitw=true';
    exportUrl += '&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=true&fzr=false';
    // margins in inches (10pt = 10/72 in ≈ 0.1389)
    exportUrl += '&top_margin=0.1389&bottom_margin=0.1389&left_margin=0.1389&right_margin=0.1389';
    const token = ScriptApp.getOAuthToken();
    const resp = UrlFetchApp.fetch(exportUrl, {headers: {Authorization: 'Bearer ' + token}, muteHttpExceptions: true});
    if (resp.getResponseCode && resp.getResponseCode() !== 200) throw new Error('Gagal mengekspor PDF. Kode: ' + resp.getResponseCode());
    const blob = resp.getBlob();
    return Utilities.base64Encode(blob.getBytes());
  } catch (e) {
    throw new Error('downloadJawabanEssayPDF error: ' + (e && e.message ? e.message : e));
  }
}

/**
 * Return JAWABAN_ESSAY sheet as CSV string (for client-side PDF conversion)
 */
function downloadJawabanEssayCSV() {
  const sh = ss().getSheetByName('JAWABAN_ESSAY');
  if (!sh) return 'Tidak ada sheet JAWABAN_ESSAY';
  const data = sh.getDataRange().getValues();
  if (!data || data.length === 0) return 'Tidak ada data';
  const rows = data.map(r => {
    return r.map(c => '"' + (c === undefined || c === null ? '' : ('' + c).replace(/"/g, '""')) + '"').join(',');
  });
  return rows.join('\n');
}

/**
 * Test helper to verify UrlFetchApp permission and OAuth token availability.
 * Run this from the Apps Script editor (select `testFetch` → Run).
 */
function testFetch() {
  try {
    const token = ScriptApp.getOAuthToken();
    const resp = UrlFetchApp.fetch('https://www.google.com', {headers: {Authorization: 'Bearer ' + token}, muteHttpExceptions: true});
    const code = (resp && resp.getResponseCode) ? resp.getResponseCode() : 'no-code';
    Logger.log('testFetch response code: ' + code);
    return 'OK: ' + code;
  } catch (e) {
    throw new Error('testFetch error: ' + (e && e.message ? e.message : e));
  }
}
/** 
 * ARSIP SMART CBT 2026 - BACKEND
 * Pastikan SPREADSHEET_ID sudah benar!
 */
const SPREADSHEET_ID = '1Pllmtr6nAYEJtmw2K4eVjV86tdCK2tg-kHTUmlQ-mI4'; // <-- GANTI DENGAN ID SPREADSHEET ANDA

function ss() {
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (e) {
    throw new Error('Gagal membuka spreadsheet. Periksa SPREADSHEET_ID dan izin: ' + e.message);
  }
}

/**
 * Catat informasi login sederhana ke sheet LOG_LOGIN (TIMESTAMP, NAMA, KELAS, ABSEN)
 * Dipanggil dari frontend untuk merekam kehadiran siswa/admin.
 */
function simpanLogin(nama, kelas, absen, asalSekolah) {
  try {
    const shName = 'LOG_LOGIN';
    const ssObj = ss();
    let sh = ssObj.getSheetByName(shName);
    if (!sh) {
      sh = ssObj.insertSheet(shName);
      sh.appendRow(['TIMESTAMP','NAMA','KELAS','ABSEN','NPSN']);
    }
    const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    sh.appendRow([ts, nama || '', kelas || '', absen || '', asalSekolah || '']);
    return 'OK';
  } catch (e) {
    // jangan lempar error yang memutus frontend — kembalikan pesan kegagalan
    return 'ERR:' + (e && e.message ? e.message : e);
  }
}

/**
 * Get list of unique schools from LOG_LOGIN sheet
 */
function getSekolahList() {
  try {
    const npsns = new Set();
    
    // Try to get NPSN from LOG_LOGIN sheet first
    try {
      const shLogin = ss().getSheetByName('LOG_LOGIN');
      if (shLogin) {
        const dataLogin = shLogin.getDataRange().getValues();
        if (dataLogin && dataLogin.length > 1) {
          const headersLogin = dataLogin[0].map(h => (h || '').toString().toUpperCase().replace(/\s+/g, ''));
          let npsninxLogin = headersLogin.indexOf('NPSN');
          if (npsninxLogin === -1) npsninxLogin = headersLogin.indexOf('ASAL_SEKOLAH');
          
          if (npsninxLogin !== -1) {
            for (let i = 1; i < dataLogin.length; i++) {
              const npsn = (dataLogin[i][npsninxLogin] || '').toString().trim();
              if (npsn) npsns.add(npsn);
            }
          }
        }
      }
    } catch (eLogin) {
      Logger.log('getSekolahList: error reading LOG_LOGIN: ' + eLogin);
    }
    
    // Also get NPSN from JAWABAN_ESSAY sheet as fallback/supplement
    try {
      const shEssay = ss().getSheetByName('JAWABAN_ESSAY');
      if (shEssay) {
        const dataEssay = shEssay.getDataRange().getValues();
        if (dataEssay && dataEssay.length > 1) {
          const headersEssay = dataEssay[0].map(h => (h || '').toString().toUpperCase().replace(/\s+/g, ''));
          Logger.log('getSekolahList: JAWABAN_ESSAY headers = ' + JSON.stringify(headersEssay));
          
          let npsninxEssay = headersEssay.indexOf('NPSN');
          if (npsninxEssay === -1) npsninxEssay = headersEssay.indexOf('ASAL_SEKOLAH');
          
          Logger.log('getSekolahList: NPSN index in JAWABAN_ESSAY = ' + npsninxEssay);
          
          if (npsninxEssay !== -1) {
            for (let i = 1; i < dataEssay.length; i++) {
              const npsn = (dataEssay[i][npsninxEssay] || '').toString().trim();
              if (npsn) {
                Logger.log('getSekolahList: found NPSN = ' + npsn);
                npsns.add(npsn);
              }
            }
          } else {
            Logger.log('getSekolahList: NPSN column not found in JAWABAN_ESSAY');
          }
        }
      } else {
        Logger.log('getSekolahList: JAWABAN_ESSAY sheet not found');
      }
    } catch (eEssay) {
      Logger.log('getSekolahList: error reading JAWABAN_ESSAY: ' + eEssay);
    }
    
    // Also try HASIL sheet
    try {
      const shHasil = ss().getSheetByName('HASIL');
      if (shHasil) {
        const dataHasil = shHasil.getDataRange().getValues();
        if (dataHasil && dataHasil.length > 1) {
          const headersHasil = dataHasil[0].map(h => (h || '').toString().toUpperCase().replace(/\s+/g, ''));
          let npsninxHasil = headersHasil.indexOf('NPSN');
          
          if (npsninxHasil !== -1) {
            for (let i = 1; i < dataHasil.length; i++) {
              const npsn = (dataHasil[i][npsninxHasil] || '').toString().trim();
              if (npsn) npsns.add(npsn);
            }
          }
        }
      }
    } catch (eHasil) {
      Logger.log('getSekolahList: error reading HASIL: ' + eHasil);
    }
    
    Logger.log('getSekolahList: final npsns = ' + JSON.stringify(Array.from(npsns)));
    // Return sorted array
    return Array.from(npsns).sort();
  } catch (e) {
    return [];
  }
}

/**
 * List all SOAL_ sheets as objects { kelas, mapel }.
 * Used by frontend to populate class/mapel pickers.
 */
function listSoalSheets() {
  const sheets = ss().getSheets();
  const out = [];
  sheets.forEach(sh => {
    const n = sh.getName();
    if (n && n.toString().indexOf('SOAL_') === 0) {
      const parts = n.toString().split('_').slice(1);
      if (parts.length >= 1) {
        const kelas = parts[0] || '';
        const mapel = parts.slice(1).join('_') || '';
        out.push({ kelas: kelas, mapel: mapel, sheetName: n.toString() });
      }
    }
  });
  return out;
}

/**
 * Return array of question numbers (first column) for given kelas+mapel
 */
function getSoalNumbers(kelas, mapel) {
  if (!kelas || !mapel) return [];
  const name = 'SOAL_' + kelas.toString().toUpperCase().trim() + '_' + mapel.toString().toUpperCase().trim();
  const sh = ss().getSheetByName(name);
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    const val = data[i][0];
    if (val !== undefined && val !== null && ('' + val).toString().trim() !== '') out.push(('' + val).toString());
  }
  return out;
}

/**
 * Set image link for a specific soal (store link in column L, insert image formula in column D)
 */
function setSoalImage(kelas, mapel, nomor, link) {
  if (!kelas || !mapel || !nomor) throw new Error('Parameter kelas, mapel, nomor diperlukan');
  const name = 'SOAL_' + kelas.toString().toUpperCase().trim() + '_' + mapel.toString().toUpperCase().trim();
  const sh = ss().getSheetByName(name);
  if (!sh) throw new Error('Sheet soal tidak ditemukan: ' + name);
  const data = sh.getDataRange().getValues();
  // find row where column A equals nomor
  let found = -1;
  for (let i = 1; i < data.length; i++) {
    const v = (data[i][0] || '').toString().trim();
    if (v === ('' + nomor).toString().trim()) { found = i + 1; break; }
  }
  if (found === -1) throw new Error('Nomor soal tidak ditemukan: ' + nomor);
  try {
    // Normalize and store link into column L
    var storeLink = '';
    if (link && link.toString().trim()) {
      try { storeLink = getDirectLink(link.toString().trim()); } catch(e) { storeLink = link.toString().trim(); }
      sh.getRange(found, 12).setValue(storeLink);
      // set IMAGE formula in column D that references column L so it updates automatically
      try {
        // use IMAGE mode 4 (custom size) to force thumbnail display
        var formula = '=IF(L' + found + '="","",IMAGE(L' + found + ',4,120,90))';
        sh.getRange(found, 4).setFormula(formula);
      } catch(e) {
        // fallback: write direct IMAGE formula into column D with custom size
        try { sh.getRange(found, 4).setFormula('=IMAGE("' + storeLink.replace(/"/g, '\\"') + '",4,120,90)'); } catch(e) {}
      }
    } else {
      sh.getRange(found, 12).setValue('');
      try { sh.getRange(found, 4).clearContent(); } catch(e) {}
    }
    return 'OK';
  } catch (e) {
    throw new Error('Gagal menyimpan gambar: ' + e.message);
  }
}


// Standard header for SOAL sheets used by the app
const SOAL_TEMPLATE_HEADER = ['NO','PERTANYAAN','TIPE','GAMBAR','OPSI A','OPSI B','OPSI C','OPSI D','KUNCI','KATA KUNCI','SKOR','LINK GAMBAR'];

/**
 * Create or update a SOAL_<KELAS>_<MAPEL> sheet with the standard header.
 * If the sheet exists and first cell isn't 'NO', the header will be inserted above existing data.
 */
function createSoalTemplate(kelas, mapel) {
  if (!kelas || !mapel) throw new Error('kelas dan mapel diperlukan');
  const name = 'SOAL_' + kelas.toString().toUpperCase().trim() + '_' + mapel.toString().toUpperCase().trim();
  let sh = ss().getSheetByName(name);
  if (!sh) {
    sh = ss().insertSheet(name);
    sh.getRange(1,1,1,SOAL_TEMPLATE_HEADER.length).setValues([SOAL_TEMPLATE_HEADER]);
    return 'CREATED:' + name;
  }
  try {
    const first = (sh.getRange(1,1).getValue() || '').toString().trim().toUpperCase();
    if (first !== 'NO') {
      // insert a header row at top without destroying existing data
      sh.insertRows(1);
      sh.getRange(1,1,1,SOAL_TEMPLATE_HEADER.length).setValues([SOAL_TEMPLATE_HEADER]);
      return 'HEADER_INSERTED:' + name;
    } else {
      // ensure header columns match template (overwrite first row upto template length)
      sh.getRange(1,1,1,SOAL_TEMPLATE_HEADER.length).setValues([SOAL_TEMPLATE_HEADER]);
      return 'HEADER_UPDATED:' + name;
    }
  } catch (e) {
    throw new Error('Gagal membuat/menetapkan header SOAL: ' + e.message);
  }
}

/**
 * Apply the standard SOAL header to all existing SOAL_ sheets.
 * Returns an object with counts { total: N, updated: M }.
 */
function applyTemplateToAllSoalSheets() {
  const sheets = ss().getSheets();
  let total = 0, updated = 0;
  sheets.forEach(sh => {
    const n = sh.getName();
    if (n && n.toString().indexOf('SOAL_') === 0) {
      total++;
      try {
        const first = (sh.getRange(1,1).getValue()||'').toString().trim().toUpperCase();
        if (first !== 'NO') {
          sh.insertRows(1);
        }
        sh.getRange(1,1,1,SOAL_TEMPLATE_HEADER.length).setValues([SOAL_TEMPLATE_HEADER]);
        updated++;
      } catch (e) {
        // ignore per-sheet errors
      }
    }
  });
  return { total: total, updated: updated };
}

/**
 * Simple key/value settings stored in sheet `SETTINGS`.
 * setSetting(key, value) -> 'OK'
 * getSetting(key) -> value or empty string
 */
function setSetting(key, value) {
  if (!key) throw new Error('Key required');
  // Prefer storing settings in Script Properties for privacy.
  try {
    PropertiesService.getScriptProperties().setProperty(key, value);
  } catch (e) {
    // ignore if PropertiesService not available
  }
  // Keep a sheet-based fallback for compatibility/visibility.
  const ssObj = ss();
  let sh = ssObj.getSheetByName('SETTINGS');
  if (!sh) {
    sh = ssObj.insertSheet('SETTINGS');
    try { sh.getRange(1,1,1,2).setValues([['KEY','VALUE']]); } catch(e) {}
  }
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toString() === key) {
      sh.getRange(i+1, 2).setValue(value);
      return 'OK';
    }
  }
  sh.appendRow([key, value]);
  return 'OK';
}

function getSetting(key) {
  if (!key) return '';
  // Try Script Properties first (more private), then fallback to sheet.
  try {
    const v = PropertiesService.getScriptProperties().getProperty(key);
    if (v !== null && typeof v !== 'undefined') return v.toString();
  } catch (e) {
    // ignore and fallback to sheet
  }
  const sh = ss().getSheetByName('SETTINGS');
  if (!sh) return '';
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toString() === key) return (data[i][1] || '').toString();
  }
  return '';
}

/**
 * Return a CSV template for soal (header + example rows) as plain text.
/**
 * Fix image formulas for a given sheet name (write IMAGE() into column D for rows
 * that have a LINK GAMBAR (col L / index 12) but no IMAGE formula in column D.
 * Returns the number of rows updated.
 */
function fixSoalImagesForSheet(sheetName) {
  if (!sheetName) throw new Error('sheetName diperlukan');
  const sh = ss().getSheetByName(sheetName);
  if (!sh) throw new Error('Sheet tidak ditemukan: ' + sheetName);
  const data = sh.getDataRange().getValues();
  if (!data || data.length <= 1) return 0;
  let updated = 0;
  // iterate rows starting from 2 (1-based)
  for (let i = 1; i < data.length; i++) {
    const rowIndex = i + 1;
    try {
      const link = (data[i][11] || '').toString().trim(); // col L -> index 11 (0-based)
      if (!link) continue;
      const cell = sh.getRange(rowIndex, 4); // col D
      const existingFormula = cell.getFormula ? cell.getFormula() : '';
      if (existingFormula && existingFormula.toString().toUpperCase().indexOf('IMAGE(') !== -1) continue;
      // sanitize link
      const safe = link.replace(/"/g, '"').replace(/\r?\n/g, ' ').trim();
      // set as thumbnail (mode 4) with width=120px and height=90px
      cell.setFormula('=IMAGE("' + safe + '",4,120,90)');
      updated++;
    } catch (e) {
      // skip problematic rows
    }
  }
  return updated;
}

/**
 * Fix image formulas for a kelas+mapel (builds sheet name `SOAL_<KELAS>_<MAPEL>`).
 */
function fixSoalImages(kelas, mapel) {
  if (!kelas || !mapel) throw new Error('Parameter kelas dan mapel diperlukan');
  const name = 'SOAL_' + kelas.toString().toUpperCase().trim() + '_' + mapel.toString().toUpperCase().trim();
  return fixSoalImagesForSheet(name);
}

/**
 * Iterate all sheets starting with `SOAL_` and fix image formulas where needed.
 * Returns summary object { totalSheets: N, totalFixed: M, details: [{sheet, fixed}] }
 */
function fixAllSoalImages() {
  const sheets = ss().getSheets();
  const details = [];
  let totalFixed = 0;
  let totalSheets = 0;
  sheets.forEach(sh => {
    const n = sh.getName();
    if (n && n.toString().indexOf('SOAL_') === 0) {
      totalSheets++;
      try {
        const fixed = fixSoalImagesForSheet(n);
        details.push({ sheet: n, fixed: fixed });
        totalFixed += fixed;
      } catch (e) {
        details.push({ sheet: n, error: (e && e.message) ? e.message : String(e) });
      }
    }
  });
  return { totalSheets: totalSheets, totalFixed: totalFixed, details: details };
}
 
/**
 * Called from client: google.script.run.getTemplate()
 */
function getTemplate() {
  try {
    // Return tab-separated template matching the TEMPLATE sheet header
    var headerLine = 'NO\tPERTANYAAN\tTIPE\tGAMBAR\tOPSI A\tOPSI B\tOPSI C\tOPSI D\tKUNCI\tKATA KUNCI\tSKOR\tLINK GAMBAR';
    var pgRow = ['1','Contoh soal pilihan ganda','PG','','Pilihan A','Pilihan B','Pilihan C','Pilihan D','A','', '4',''].join('\t');
    var essayRow = ['2','Contoh soal essay: jelaskan...','ESSAY','','','','','','','kata_kunci1,kata_kunci2','10',''].join('\t');
    return headerLine + '\n' + pgRow + '\n' + essayRow;
  } catch (e) {
    throw new Error('getTemplate error: ' + (e && e.message ? e.message : e));
  }
}

/**
 * Migrate all key/value pairs from sheet `SETTINGS` into Script Properties.
 * Overwrites any existing script properties with the same keys.
 * Returns summary string 'MIGRATED:<count>' or message when no sheet.
 */
function migrateSettingsToProperties() {
  const sh = ss().getSheetByName('SETTINGS');
  if (!sh) return 'No SETTINGS sheet found';
  const data = sh.getDataRange().getValues();
  if (!data || data.length <= 1) return 'No settings to migrate';
  let count = 0;
  try {
    const props = PropertiesService.getScriptProperties();
    for (let i = 1; i < data.length; i++) {
      const k = (data[i][0] || '').toString().trim();
      const v = (data[i][1] || '').toString();
      if (!k) continue;
      props.setProperty(k, v);
      count++;
    }
    return 'MIGRATED:' + count;
  } catch (e) {
    throw new Error('Migration failed: ' + e.message);
  }
}


function listSoal(userKelas) {
  // Gabungkan soal dari FORM_LINKS dan SOAL_CBT
  const isGuru = (userKelas === 'GURU_ALL');
  const normalizedUserKelas = (userKelas || '').toString().trim().toUpperCase();
  const result = [];
  const formSheet = ss().getSheetByName('FORM_LINKS');
  const cbtSheet = ss().getSheetByName('SOAL_CBT');
  let formData = [], cbtData = [];
  if (formSheet && formSheet.getLastRow() > 1) {
    formData = formSheet.getDataRange().getValues().slice(1); // skip header
  }
  if (cbtSheet && cbtSheet.getLastRow() > 1) {
    cbtData = cbtSheet.getDataRange().getValues().slice(1); // skip header
  }
  // Buat map untuk lookup cepat
  const formMap = {};
  formData.forEach(row => {
    const kelas = (row[0] || '').toString().trim();
    const mapel = (row[1] || '').toString().trim();
    if (!kelas || !mapel) return;
    const key = kelas.toUpperCase() + '___' + mapel.toUpperCase();
    formMap[key] = row;
  });
  const cbtMap = {};
  cbtData.forEach(row => {
    const kelas = (row[0] || '').toString().trim();
    const mapel = (row[1] || '').toString().trim();
    if (!kelas || !mapel) return;
    const key = kelas.toUpperCase() + '___' + mapel.toUpperCase();
    cbtMap[key] = row;
  });
  // Gabungkan semua key unik
  const allKeys = new Set([...Object.keys(formMap), ...Object.keys(cbtMap)]);
  allKeys.forEach(key => {
    const formRow = formMap[key];
    const cbtRow = cbtMap[key];
    const kelas = key.split('___')[0];
    const mapel = key.split('___')[1];
    // Cek kecocokan kelas
    const normalizedFormKelas = kelas.toUpperCase();
    const kelasMatch = normalizedFormKelas === normalizedUserKelas
      || normalizedUserKelas.startsWith(normalizedFormKelas)
      || normalizedFormKelas.startsWith(normalizedUserKelas);
    if (!isGuru && !kelasMatch) return;
    // Gabungan CBT+FORM
    if (formRow && cbtRow) {
      const formStatus = (formRow[4] || 'AKTIF').toString().trim().toUpperCase();
      const cbtStatus = (cbtRow[2] || 'AKTIF').toString().trim().toUpperCase();
      // determine upload timestamp from both sources (prefer latest)
      let uploadTs = 0;
      try {
        if (formRow[3]) {
          const d1 = new Date(formRow[3]);
          if (!isNaN(d1.getTime())) uploadTs = Math.max(uploadTs, d1.getTime());
        }
      } catch (e) { }
      try {
        if (cbtRow[3]) {
          const d2 = new Date(cbtRow[3]);
          if (!isNaN(d2.getTime())) uploadTs = Math.max(uploadTs, d2.getTime());
        }
      } catch (e) { }
      result.push({
        KELAS: kelas,
        MAPEL: mapel,
        STATUS: (cbtStatus === 'AKTIF' || formStatus === 'AKTIF') ? 'AKTIF' : 'NONAKTIF',
        FORM_STATUS: formStatus,
        CBT_STATUS: cbtStatus,
        TYPE: 'BOTH',
        FORM_LINK: formRow[2] || '',
        UPLOAD_TS: uploadTs
      });
    } else if (formRow) {
      const formStatus = (formRow[4] || 'AKTIF').toString().trim().toUpperCase();
      // read upload timestamp from FORM_LINKS (col index 3)
      let uploadTs = 0;
      try {
        if (formRow[3]) {
          const d = new Date(formRow[3]);
          if (!isNaN(d.getTime())) uploadTs = d.getTime();
        }
      } catch (e) { uploadTs = 0; }
      result.push({
        KELAS: kelas,
        MAPEL: mapel,
        STATUS: formStatus,
        FORM_STATUS: formStatus,
        TYPE: 'FORM',
        FORM_LINK: formRow[2] || '',
        UPLOAD_TS: uploadTs
      });
    } else if (cbtRow) {
      const cbtStatus = (cbtRow[2] || 'AKTIF').toString().trim().toUpperCase();
      // try to read an upload timestamp from CBT row (if present in a later column)
      let uploadedTs = 0;
      try {
        if (cbtRow[3]) {
          const d = new Date(cbtRow[3]);
          if (!isNaN(d.getTime())) uploadedTs = d.getTime();
        }
      } catch (e) { uploadedTs = 0; }
      result.push({
        KELAS: kelas,
        MAPEL: mapel,
        STATUS: cbtStatus,
        CBT_STATUS: cbtStatus,
        TYPE: 'CBT',
        FORM_LINK: '',
        UPLOAD_TS: uploadedTs
      });
    }
  });
  // Sort result:
  // 1) Items with upload timestamp first (newest -> oldest)
  // 2) Then prioritize kelas starting with 7,8,9 in that order
  // 3) Then numeric kelas ascending
  // 4) Then MAPEL alphabetically
  result.sort(function(a, b) {
    const ta = (a.UPLOAD_TS || 0);
    const tb = (b.UPLOAD_TS || 0);
    if (ta !== tb) return tb - ta; // newer uploads first

    function extractNum(k) {
      const m = (k || '').toString().match(/\d+/);
      return m ? parseInt(m[0], 10) : 999;
    }
    function pref(n) { if (n === 7) return 0; if (n === 8) return 1; if (n === 9) return 2; return 3; }
    const na = extractNum(a.KELAS);
    const nb = extractNum(b.KELAS);
    const pa = pref(na);
    const pb = pref(nb);
    if (pa !== pb) return pa - pb;
    if (na !== nb) return na - nb;
    const mcmp = (a.MAPEL || '').toString().localeCompare((b.MAPEL || '').toString());
    if (mcmp !== 0) return mcmp;
    return (a.KELAS || '').toString().localeCompare((b.KELAS || '').toString());
  });

  return result;
}

function getSoal(kelas, mapel) {
  if (!kelas || !mapel) throw new Error('Parameter `kelas` dan `mapel` wajib diisi');
  const name = 'SOAL_' + kelas.toString().toUpperCase().trim() + '_' + mapel.toString().toUpperCase().trim();
  const sh = ss().getSheetByName(name);
  if (!sh) throw new Error('Sheet soal tidak ditemukan: ' + name);

  const data = sh.getDataRange().getValues();
  if (!data || data.length < 2) return { mapel: mapel, durasi: 90, data: [] };

  const headers = data[0];
  const rows = data.slice(1);


  const norm = v => (v || '').toString().toUpperCase().replace(/\s+/g, '').replace(/[^A-Z0-9]/g, '');
  const getIdx = (names) => {
    const normalizedHeaders = headers.map(h => norm(h));
    const targetNames = names.map(n => norm(n));
    for (let target of targetNames) {
      const idx = normalizedHeaders.indexOf(target);
      if (idx !== -1) return idx;
    }
    return -1;
  };

  const idxPertanyaan = getIdx(['PERTANYAAN', 'SOAL', 'QUESTION', 'ISI_SOAL', 'TEKS_SOAL']);
  let idxA = getIdx(['OPSI_A', 'A', 'OPSI A', 'JAWABAN_A', 'PILIHAN_A']);
  let idxB = getIdx(['OPSI_B', 'B', 'OPSI B', 'JAWABAN_B', 'PILIHAN_B']);
  let idxC = getIdx(['OPSI_C', 'C', 'OPSI C', 'JAWABAN_C', 'PILIHAN_C']);
  let idxD = getIdx(['OPSI_D', 'D', 'OPSI D', 'JAWABAN_D', 'PILIHAN_D']);
  let idxGambar = getIdx(['GAMBAR','IMAGE','GAMBAR_URL']);
  let idxKunci = getIdx(['KUNCI', 'KEY', 'JAWABAN', 'KUNCI_JAWABAN']);
  let idxTipe = getIdx(['TIPE', 'TYPE']);
  // also detect LINK GAMBAR column (backup URL stored in sheet)
  let idxLinkGambar = getIdx(['LINK GAMBAR', 'LINK_GAMBAR', 'LINK', 'LINKIMAGE']);

  if (idxPertanyaan !== -1 && (idxA === -1 || idxB === -1 || idxC === -1 || idxD === -1)) {
    idxA = idxA === -1 ? idxPertanyaan + 1 : idxA;
    idxB = idxB === -1 ? idxPertanyaan + 2 : idxB;
    idxC = idxC === -1 ? idxPertanyaan + 3 : idxC;
    idxD = idxD === -1 ? idxPertanyaan + 4 : idxD;
  }
  if (idxKunci === -1 && rows.length && rows[0].length > 0) {
    idxKunci = rows[0].length - 1;
  }

  const safeGet = (row, idx) => (idx >= 0 && idx < row.length) ? row[idx] : '';

  const formattedData = rows
    .map(function(r, rowIdx) {
      let tipe = '';
      if (typeof r[idxTipe] !== 'undefined') tipe = r[idxTipe];
      tipe = (tipe || '').toString().trim().toUpperCase();
      // prefer LINK GAMBAR (col L) first, otherwise fallback to GAMBAR column
      let gambarVal = (safeGet(r, idxLinkGambar) || safeGet(r, idxGambar));
      // If the stored value is a display text (eg. "Link Gambar") because the sheet cell contains a HYPERLINK formula,
      // try to read the actual formula and extract the URL from HYPERLINK(...) or IMAGE(...)
      try {
        const sheetRow = rowIdx + 2; // rows[] is data.slice(1) so rowIdx 0 -> sheet row 2
        if ((!gambarVal || !(/^https?:\/\//i.test(gambarVal) || /^data:/i.test(gambarVal))) && idxLinkGambar !== -1) {
          const f = sh.getRange(sheetRow, idxLinkGambar + 1).getFormula() || '';
          if (f) {
            // try HYPERLINK("url","text")
            const m = f.match(/HYPERLINK\s*\(\s*"([^"]+)"/i);
            if (m && m[1]) gambarVal = m[1];
            else {
              // try IMAGE("url")
              const m2 = f.match(/IMAGE\s*\(\s*"([^"]+)"/i);
              if (m2 && m2[1]) gambarVal = m2[1];
            }
          }
        }
      } catch (e) { /* ignore */ }
      return {
        soal: safeGet(r, idxPertanyaan),
        gambar: gambarVal,
        opsi_a: safeGet(r, idxA),
        opsi_b: safeGet(r, idxB),
        opsi_c: safeGet(r, idxC),
        opsi_d: safeGet(r, idxD),
        kunci: safeGet(r, idxKunci).toString().trim().toUpperCase(),
        tipe: tipe
      };
    })
    .filter(item => item.soal);

  // Only shuffle PG questions (if sheet-level 'ACAK' flag enabled).
  // Always present PG items first, then ESSAY items in original order.
  function getAcakFlag(kelasName, mapelName) {
    try {
      const cbt = ss().getSheetByName('SOAL_CBT');
      if (cbt) {
        const d = cbt.getDataRange().getValues();
        for (let r = 1; r < d.length; r++) {
          const rk = (d[r][0] || '').toString().trim().toUpperCase();
          const rm = (d[r][1] || '').toString().trim().toUpperCase();
          if (rk === kelasName.toString().trim().toUpperCase() && rm === mapelName.toString().trim().toUpperCase()) {
            const val = (d[r][4] || '').toString().trim().toUpperCase();
            return (val === 'YA' || val === '1' || val === 'TRUE');
          }
        }
      }
      // fallback: read Z2 marker on sheet
      const sheetName = 'SOAL_' + kelasName.toString().trim().toUpperCase() + '_' + mapelName.toString().trim().toUpperCase();
      const per = ss().getSheetByName(sheetName);
      if (per) {
        try {
          const z2 = (per.getRange('Z2').getValue() || '').toString().trim().toUpperCase();
          return (z2 === 'YA' || z2 === '1' || z2 === 'TRUE');
        } catch (e) { /* ignore */ }
      }
    } catch (e) { /* ignore */ }
    return false;
  }

  const pgItems = [];
  const essayItems = [];
  formattedData.forEach(item => {
    const t = (item.tipe || '').toString().trim().toUpperCase();
    if (t === 'ESSAY') essayItems.push(item); else pgItems.push(item);
  });

  // determine if this sheet has acak enabled
  const acakEnabled = getAcakFlag(kelas, mapel);
  if (acakEnabled) {
    // Fisher-Yates shuffle PG items only
    for (let i = pgItems.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      const tmp = pgItems[i]; pgItems[i] = pgItems[j]; pgItems[j] = tmp;
    }
  }

  const finalData = pgItems.concat(essayItems);

  return {
    mapel: mapel,
    durasi: 90,
    data: finalData
  };
}

function uploadSoalCSV(mapel, kelas, csvContent) {
  if (typeof csvContent !== 'string' || !csvContent.trim()) {
    throw new Error("Gagal membaca konten file CSV. Pastikan Anda mengunggah file teks.");
  }
  const name = 'SOAL_' + kelas.toUpperCase().trim() + '_' + mapel.toUpperCase().trim();
  const dataSoal = Utilities.parseCsv(csvContent);
  if (!dataSoal || dataSoal.length === 0) { throw new Error("File CSV kosong atau tidak valid."); }

  let sh = ss().getSheetByName(name);
  if (!sh) sh = ss().insertSheet(name);
  else sh.clear();
  sh.getRange(1, 1, dataSoal.length, dataSoal[0].length).setValues(dataSoal);
  // If CSV did not include the expected header (first cell != NO), insert the standard header
  try {
    const firstCell = (sh.getRange(1,1).getValue()||'').toString().trim().toUpperCase();
    if (firstCell !== 'NO') {
      sh.insertRows(1);
      sh.getRange(1,1,1,SOAL_TEMPLATE_HEADER.length).setValues([SOAL_TEMPLATE_HEADER]);
    } else {
      // ensure header columns conform to template (overwrite up to template length)
      sh.getRange(1,1,1,SOAL_TEMPLATE_HEADER.length).setValues([SOAL_TEMPLATE_HEADER]);
    }
  } catch (e) {
    // ignore header adjustment errors
  }
  try { sh.getRange("Z1").setValue("AKTIF"); } catch (e) {}

  // Tambahkan/Update ke sheet SOAL_CBT
  let cbtSheet = ss().getSheetByName('SOAL_CBT');
  if (!cbtSheet) {
    cbtSheet = ss().insertSheet('SOAL_CBT');
    cbtSheet.appendRow(['KELAS', 'MAPEL', 'STATUS', 'TANGGAL_UPLOAD']);
  }
  const normalizedKelas = kelas.toUpperCase().trim();
  const normalizedMapel = mapel.toUpperCase().trim();
  const cbtData = cbtSheet.getDataRange().getValues();
  let found = false;
  for (let i = 1; i < cbtData.length; i++) {
    if ((cbtData[i][0] || '').toString().toUpperCase().trim() === normalizedKelas && (cbtData[i][1] || '').toString().toUpperCase().trim() === normalizedMapel) {
      cbtSheet.getRange(i + 1, 3).setValue('AKTIF');
      try { cbtSheet.getRange(i + 1, 4).setValue(new Date()); } catch (e) {}
      found = true;
      break;
    }
  }
  if (!found) {
    cbtSheet.appendRow([normalizedKelas, normalizedMapel, 'AKTIF', new Date()]);
  }

  // After uploading CSV, attempt to set IMAGE formulas for any links in column L on this new sheet
  try {
    var fixedCount = 0;
    try { fixedCount = fixSoalImagesForSheet(name) || 0; } catch(e) { fixedCount = 0; }
    return 'Berhasil: ' + name + ' Aktif. Gambar diproses: ' + fixedCount + ' baris.';
  } catch (e) {
    return 'Berhasil: ' + name + ' Aktif.';
  }
}

/**
 * Save or update a Google Form link into sheet `FORM_LINKS`.
 * Columns: KELAS (0), MAPEL (1), LINK (2), TANGGAL_UPLOAD (3), STATUS (4)
 */
function uploadLinkGoogleForm(kelas, mapel, link) {
  if (!kelas || !mapel || !link) throw new Error('Parameter kelas, mapel, link diperlukan');
  kelas = kelas.toString().toUpperCase().trim();
  mapel = mapel.toString().toUpperCase().trim();
  link = link.toString().trim();
  // Normalize common Google Form links into an embeddable view URL
  try {
    var l = link;
    if (!/^https?:\/\//i.test(l)) l = 'https://' + l;
    if (l.indexOf('docs.google.com/forms') !== -1) {
      // try to extract form id from /d/e/<id>/ or /d/<id>/ patterns
      var m = l.match(/\/d\/e\/([a-zA-Z0-9_-]+)/) || l.match(/\/d\/([a-zA-Z0-9_-]+)/);
      if (m && m[1]) {
        l = 'https://docs.google.com/forms/d/e/' + m[1] + '/viewform?embedded=true';
      } else {
        // ensure viewform + embedded flag
        if (l.indexOf('viewform') === -1) {
          if (l.endsWith('/')) l = l + 'viewform'; else l = l + (l.indexOf('?') === -1 ? '?viewform' : '&viewform');
        }
        if (l.indexOf('embedded=true') === -1) l = l + (l.indexOf('?') === -1 ? '?embedded=true' : '&embedded=true');
      }
    }
    link = l;
  } catch (e) {
    // leave original link if normalization fails
  }
  const ssObj = ss();
  let sh = ssObj.getSheetByName('FORM_LINKS');
  if (!sh) {
    sh = ssObj.insertSheet('FORM_LINKS');
    try { sh.appendRow(['KELAS','MAPEL','LINK','TANGGAL_UPLOAD','STATUS']); } catch(e) {}
  }
  const data = sh.getDataRange().getValues();
  // try to find existing row
  for (let i = 1; i < data.length; i++) {
    const rK = (data[i][0] || '').toString().toUpperCase().trim();
    const rM = (data[i][1] || '').toString().toUpperCase().trim();
    if (rK === kelas && rM === mapel) {
      try {
        sh.getRange(i+1, 3).setValue(link);
        sh.getRange(i+1, 4).setValue(new Date());
        sh.getRange(i+1, 5).setValue('AKTIF');
      } catch (e) { throw new Error('Gagal memperbarui FORM_LINKS: ' + e.message); }
      return 'UPDATED';
    }
  }
  // append new row
  try {
    sh.appendRow([kelas, mapel, link, new Date(), 'AKTIF']);
  } catch (e) { throw new Error('Gagal menambahkan ke FORM_LINKS: ' + e.message); }
  return 'CREATED';
}

function toggleSoalStatus(kelas, mapel) {
  // Toggle status in SOAL_CBT (kolom C) for the given kelas+mapel.
  const ssObj = ss();
  const cbtSheetName = 'SOAL_CBT';
  let cbt = ssObj.getSheetByName(cbtSheetName);
  if (!cbt) {
    // create sheet with header if missing
    cbt = ssObj.insertSheet(cbtSheetName);
    try { cbt.appendRow(['KELAS', 'MAPEL', 'STATUS']); } catch (e) { }
  }

  const data = cbt.getDataRange().getValues();
  const normalizedKelas = (kelas || '').toString().toUpperCase().trim();
  const normalizedMapel = (mapel || '').toString().toUpperCase().trim();
  let foundRow = -1;
  for (let i = 1; i < data.length; i++) {
    const rKelas = (data[i][0] || '').toString().toUpperCase().trim();
    const rMapel = (data[i][1] || '').toString().toUpperCase().trim();
    if (rKelas === normalizedKelas && rMapel === normalizedMapel) { foundRow = i + 1; break; }
  }

  let nextStatus = 'AKTIF';
  if (foundRow !== -1) {
    const current = (cbt.getRange(foundRow, 3).getValue() || '').toString().trim().toUpperCase();
    nextStatus = (current === 'NONAKTIF') ? 'AKTIF' : 'NONAKTIF';
    cbt.getRange(foundRow, 3).setValue(nextStatus);
  } else {
    // If not found, append a new row with AKTIF status
    try {
      cbt.appendRow([normalizedKelas, normalizedMapel, nextStatus]);
      foundRow = cbt.getLastRow();
    } catch (e) {
      // fallback: return error message
      return 'Gagal memperbarui SOAL_CBT: ' + e.message;
    }
  }

  // Also try to keep sheet-level marker in SOAL_<kelas>_<mapel> (Z1) in sync if exists
  try {
    const perSheet = ssObj.getSheetByName('SOAL_' + kelas + '_' + mapel);
    if (perSheet) {
      perSheet.getRange('Z1').setValue(nextStatus);
    }
  } catch (e) {
    // ignore errors here
  }

  return nextStatus;
}

/**
 * Set 'acak' (randomize) flag for soal per mapel and optional kelas.
 * @param {string} mapel - nama mapel (required)
 * @param {string} kelas - subkelas seperti '7A' atau base '7' (optional, empty means apply to all kelas for that mapel)
 * @param {boolean|number|string} acak - true/1/'1' to enable, false/0/'0' to disable
 * @returns {string} summary
 */
function setSoalAcakState(mapel, kelas, acak) {
  if (!mapel || !mapel.toString().trim()) throw new Error('Parameter mapel diperlukan');
  const ssObj = ss();
  let cbt = ssObj.getSheetByName('SOAL_CBT');
  if (!cbt) {
    cbt = ssObj.insertSheet('SOAL_CBT');
    try { cbt.appendRow(['KELAS','MAPEL','STATUS','TANGGAL_UPLOAD','ACAK']); } catch(e) {}
  }

  const data = cbt.getDataRange().getValues();
  const targetMapel = mapel.toString().trim().toUpperCase();
  const targetKelas = kelas && kelas.toString().trim() ? kelas.toString().trim().toUpperCase() : '';
  const flag = (acak === true || acak === '1' || acak === 1 || acak === 'true' || acak === 'TRUE') ? 'YA' : 'TIDAK';

  let updated = 0;
  // Update existing rows matching mapel (and kelas if provided)
  for (let i = 1; i < data.length; i++) {
    const rK = (data[i][0] || '').toString().trim().toUpperCase();
    const rM = (data[i][1] || '').toString().trim().toUpperCase();
    if (rM === targetMapel && (targetKelas === '' || rK === targetKelas)) {
      try { cbt.getRange(i+1, 5).setValue(flag); } catch(e) {}
      try { cbt.getRange(i+1, 4).setValue(new Date()); } catch(e) {}
      updated++;
      // try to write per-sheet marker (Z2)
      try {
        const perName = 'SOAL_' + rK + '_' + rM;
        const perSheet = ssObj.getSheetByName(perName);
        if (perSheet) perSheet.getRange('Z2').setValue(flag);
      } catch(e) {}
    }
  }

  // If no existing rows updated and kelas specified, append a new row
  if (updated === 0 && targetKelas) {
    try { cbt.appendRow([targetKelas, targetMapel, 'AKTIF', new Date(), flag]); updated = 1; } catch(e) {}
    try { const perSheet = ssObj.getSheetByName('SOAL_' + targetKelas + '_' + targetMapel); if (perSheet) perSheet.getRange('Z2').setValue(flag); } catch(e) {}
  }

  // If kelas not specified, attempt to apply to all existing SOAL_ sheets matching mapel
  if (updated === 0 && !targetKelas) {
    const sheets = ssObj.getSheets();
    for (let sh of sheets) {
      const n = sh.getName();
      if (n && n.indexOf('SOAL_') === 0) {
        const parts = n.split('_').slice(1);
        if (parts.length >= 2) {
          const rK = parts[0].toString().trim().toUpperCase();
          const rM = parts.slice(1).join('_').toString().trim().toUpperCase();
          if (rM === targetMapel) {
            // ensure entry exists in SOAL_CBT
            let found = false;
            for (let i = 1; i < data.length; i++) {
              const rrK = (data[i][0] || '').toString().trim().toUpperCase();
              const rrM = (data[i][1] || '').toString().trim().toUpperCase();
              if (rrK === rK && rrM === rM) {
                try { cbt.getRange(i+1,5).setValue(flag); } catch(e) {}
                try { cbt.getRange(i+1,4).setValue(new Date()); } catch(e) {}
                found = true; updated++; break;
              }
            }
            if (!found) {
              try { cbt.appendRow([rK, rM, 'AKTIF', new Date(), flag]); updated++; } catch(e) {}
            }
            // write per-sheet marker
            try { sh.getRange('Z2').setValue(flag); } catch(e) {}
          }
        }
      }
    }
  }

  return 'UPDATED:' + updated;
}

function toggleFormLinkStatus(kelas, mapel) {
  const sh = ss().getSheetByName('FORM_LINKS');
  if (!sh) return 'Sheet FORM_LINKS tidak ditemukan';
  const data = sh.getDataRange().getValues();
  const normalizedKelas = kelas.toUpperCase().trim();
  const normalizedMapel = mapel.toUpperCase().trim();

  for (let i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toString().toUpperCase().trim() === normalizedKelas && (data[i][1] || '').toString().toUpperCase().trim() === normalizedMapel) {
      const current = (data[i][4] || '').toString().trim().toUpperCase();
      const nextStatus = current === 'NONAKTIF' ? 'AKTIF' : 'NONAKTIF';
      sh.getRange(i + 1, 5).setValue(nextStatus);
      return nextStatus;
    }
  }
  return 'Link tidak ditemukan';
}

function deleteSoal(kelas, mapel) {
  const ssObj = ss();
  const sheetName = 'SOAL_' + kelas + '_' + mapel;
  const sh = ssObj.getSheetByName(sheetName);
  let deletedAny = false;
  if (sh) {
    try { ssObj.deleteSheet(sh); deletedAny = true; } catch (e) { /* ignore */ }
  }

  // Helper to remove matching rows from a sheet (keep header)
  const removeFromSheet = (sheetNameTarget, colKelasIndex, colMapelIndex) => {
    const s = ssObj.getSheetByName(sheetNameTarget);
    if (!s) return false;
    const data = s.getDataRange().getValues();
    if (!data || data.length <= 1) return false;
    const normalizedK = (kelas || '').toString().toUpperCase().trim();
    const normalizedM = (mapel || '').toString().toUpperCase().trim();
    const keep = [data[0]];
    for (let i = 1; i < data.length; i++) {
      const rK = (data[i][colKelasIndex] || '').toString().toUpperCase().trim();
      const rM = (data[i][colMapelIndex] || '').toString().toUpperCase().trim();
      if (!(rK === normalizedK && rM === normalizedM)) keep.push(data[i]);
    }
    if (keep.length === data.length) return false;
    // determine max columns to write
    const maxCols = Math.max(...keep.map(r => r.length));
    const padded = keep.map(r => {
      const copy = r.slice();
      while (copy.length < maxCols) copy.push('');
      return copy;
    });
    s.clear();
    s.getRange(1, 1, padded.length, maxCols).setValues(padded);
    return true;
  };

  // Remove entries from SOAL_CBT (KELAS in col 0, MAPEL in col 1)
  try { if (removeFromSheet('SOAL_CBT', 0, 1)) deletedAny = true; } catch (e) { /* ignore */ }
  // Remove entries from FORM_LINKS (KELAS in col 0, MAPEL in col 1)
  try { if (removeFromSheet('FORM_LINKS', 0, 1)) deletedAny = true; } catch (e) { /* ignore */ }

  return deletedAny ? "Terhapus" : "Sheet tidak ditemukan";
}

function resetSemuaHasil() {
  const ssObj = ss();
  // Clear HASIL (keep header)
  let sh = ssObj.getSheetByName('HASIL');
  if (sh) {
    sh.clear();
    sh.appendRow(['TIMESTAMP','NO ABSEN/PESERTA','SUB KELAS','NAMA','BASE KELAS','MAPEL','BENAR','SALAH','SKOR']);
  }
  // Also clear JAWABAN_ESSAY (keep header)
  let shEssay = ssObj.getSheetByName('JAWABAN_ESSAY');
  if (shEssay) {
    shEssay.clear();
    shEssay.appendRow(['TIMESTAMP','NAMA','KELAS','ABSEN','MAPEL','NO_SOAL','SOAL','JAWABAN','SKOR']);
  }
  // Also clear LOG_LOGIN (keep header)
  let shLog = ssObj.getSheetByName('LOG_LOGIN');
  if (shLog) {
    shLog.clear();
    shLog.appendRow(['TIMESTAMP','NAMA','KELAS','ABSEN']);
  }

  return "Sheet HASIL, JAWABAN_ESSAY, dan LOG_LOGIN telah dikosongkan!";
}

/**
 * Simpan hasil akhir siswa ke sheet HASIL.
 * @param {string} mapel
 * @param {string} nama
 * @param {string} kelas
 * @param {string} absen
 * @param {number} benar
 * @param {number} salah
 * @param {number} skor  -- percentage 0..100
 */
function simpanHasil(mapel, nama, kelas, absen, benar, salah, skor, npsn) {
  if (!mapel) throw new Error('Parameter mapel diperlukan');
  const shName = 'HASIL';
  const ssObj = ss();
  let sh = ssObj.getSheetByName(shName);
  if (!sh) {
    sh = ssObj.insertSheet(shName);
    sh.appendRow(['TIMESTAMP','NO ABSEN/PESERTA','SUB KELAS','NAMA','BASE KELAS','MAPEL','BENAR','SALAH','SKOR','NPSN']);
  }
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const subkelas = (kelas || '').toString().trim();
  const base = (subkelas && subkelas.toString().match(/^([789])/)) ? subkelas.toString().match(/^([789])/)[1] : '';
  const row = [ts, absen || '', subkelas || '', nama || '', base || '', mapel || '', Number(benar) || 0, Number(salah) || 0, Number(skor) || 0, npsn || ''];

  try {
    sh.appendRow(row);
    return 'OK';
  } catch (e) {
    throw new Error('Gagal menyimpan hasil: ' + e.message);
  }
}

/**
 * Count number of rows in `HASIL` that would be deleted for given mapel and optional subkelas.
 * Returns integer count.
 */
function countHasilToDelete(mapel, subkelas, npsn) {
  if (!mapel || !mapel.toString().trim()) return 0;
  const sh = ss().getSheetByName('HASIL');
  if (!sh) return 0;
  const data = sh.getDataRange().getValues();
  if (!data || data.length <= 1) return 0;
  const targetMapel = mapel.toString().trim().toUpperCase();
  const targetSub = subkelas && subkelas.toString().trim() ? subkelas.toString().trim().toUpperCase() : '';
  const targetNpsn = npsn && npsn.toString().trim() ? npsn.toString().trim() : '';
  let cnt = 0;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const namaMapel = (row[5] || '').toString().trim().toUpperCase();
    if (namaMapel !== targetMapel) continue;
    if (targetSub) {
      const rowSub = (row[2] || '').toString().trim().toUpperCase();
      if (rowSub !== targetSub) continue;
    }
    if (targetNpsn) {
      const rowNpsn = (row[9] || '').toString().trim();
      if (rowNpsn !== targetNpsn) continue;
    }
    cnt++;
  }
  return cnt;
}

/**
 * Remove rows from `HASIL` matching given mapel, optional subkelas, and optional NPSN.
 * Returns a message with number of rows removed or an informative string.
 */
function resetHasilPerMapel(mapel, subkelas, npsn) {
  if (!mapel || !mapel.toString().trim()) return 'Parameter mapel diperlukan';
  const sh = ss().getSheetByName('HASIL');
  if (!sh) return 'Data kosong';
  const data = sh.getDataRange().getValues();
  if (!data || data.length <= 1) return 'Tidak ada data';
  const header = data[0] || [];
  const keep = [header];
  const targetMapel = mapel.toString().trim().toUpperCase();
  const targetSub = subkelas && subkelas.toString().trim() ? subkelas.toString().trim().toUpperCase() : '';
  const targetNpsn = npsn && npsn.toString().trim() ? npsn.toString().trim() : '';
  let removed = 0;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const namaMapel = (row[5] || '').toString().trim().toUpperCase();
    const rowSub = (row[2] || '').toString().trim().toUpperCase();
    const rowNpsn = (row[9] || '').toString().trim();
    
    let matches = (namaMapel === targetMapel);
    if (matches && targetSub) matches = (rowSub === targetSub);
    if (matches && targetNpsn) matches = (rowNpsn === targetNpsn);
    
    if (matches) {
      removed++;
      continue;
    }
    keep.push(row);
  }
  if (removed === 0) return 'Tidak ada data untuk kombinasi yang dipilih';
  const maxCols = Math.max.apply(null, keep.map(r => r.length));
  const padded = keep.map(r => { const copy = r.slice(); while (copy.length < maxCols) copy.push(''); return copy; });
  sh.clear();
  sh.getRange(1, 1, padded.length, maxCols).setValues(padded);
  return 'Berhasil menghapus ' + removed + ' baris untuk mapel: ' + mapel + (targetSub ? (' (Subkelas: ' + targetSub + ')') : '') + (targetNpsn ? (' (NPSN: ' + targetNpsn + ')') : '');
}

/**
 * Count and remove functions for essay answers stored in JAWABAN_ESSAY with NPSN support.
 */
function countJawabanEssayToDelete(mapel, subkelas, npsn) {
  const sh = ss().getSheetByName('JAWABAN_ESSAY');
  if (!sh) return 0;
  const data = sh.getDataRange().getValues();
  if (!data || data.length <= 1) return 0;
  const targetMapel = mapel && mapel.toString().trim() ? mapel.toString().trim().toUpperCase() : '';
  const targetSub = subkelas && subkelas.toString().trim() ? subkelas.toString().trim().toUpperCase() : '';
  const targetNpsn = npsn && npsn.toString().trim() ? npsn.toString().trim() : '';
  let cnt = 0;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowMapel = (row[4] || '').toString().trim().toUpperCase();
    const rowKelas = (row[2] || '').toString().trim().toUpperCase();
    const rowNpsn = (row[9] || '').toString().trim();
    if (targetMapel && rowMapel !== targetMapel) continue;
    if (targetSub && rowKelas !== targetSub) continue;
    if (targetNpsn && rowNpsn !== targetNpsn) continue;
    cnt++;
  }
  return cnt;
}

function resetJawabanEssayPerMapel(mapel, subkelas, npsn) {
  const sh = ss().getSheetByName('JAWABAN_ESSAY');
  if (!sh) return 'Sheet JAWABAN_ESSAY tidak ditemukan';
  const data = sh.getDataRange().getValues();
  if (!data || data.length <= 1) return 'Tidak ada data';
  const header = data[0] || [];
  const keep = [header];
  const targetMapel = mapel && mapel.toString().trim() ? mapel.toString().trim().toUpperCase() : '';
  const targetSub = subkelas && subkelas.toString().trim() ? subkelas.toString().trim().toUpperCase() : '';
  const targetNpsn = npsn && npsn.toString().trim() ? npsn.toString().trim() : '';
  let removed = 0;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowMapel = (row[4] || '').toString().trim().toUpperCase();
    const rowKelas = (row[2] || '').toString().trim().toUpperCase();
    const rowNpsn = (row[9] || '').toString().trim();
    
    let shouldRemove = true;
    if (targetMapel && rowMapel !== targetMapel) shouldRemove = false;
    else if (targetSub && rowKelas !== targetSub) shouldRemove = false;
    else if (targetNpsn && rowNpsn !== targetNpsn) shouldRemove = false;
    
    if (shouldRemove) {
      removed++;
    } else {
      keep.push(row);
    }
  }
  if (removed === 0) return 'Tidak ada data untuk kombinasi yang dipilih';
  const maxCols = Math.max.apply(null, keep.map(r => r.length));
  const padded = keep.map(r => { const copy = r.slice(); while (copy.length < maxCols) copy.push(''); return copy; });
  sh.clear();
  sh.getRange(1, 1, padded.length, maxCols).setValues(padded);
  return 'Berhasil menghapus ' + removed + ' baris pada JAWABAN_ESSAY untuk mapel: ' + (mapel || 'ALL') + (targetSub ? (' (Subkelas: ' + targetSub + ')') : '') + (targetNpsn ? (' (NPSN: ' + targetNpsn + ')') : '');
}

function downloadHasil() {
  const sh = ss().getSheetByName('HASIL');
  return sh ? sh.getDataRange().getValues().map(r => r.join(',')).join('\n') : "Data kosong";
}

function downloadHasilPerMapel(mapel) {
  const sh = ss().getSheetByName('HASIL');
  if (!sh) return "Data kosong";
  
  const data = sh.getDataRange().getValues();
  const filtered = [data[0]]; // Header
  
  for (let i = 1; i < data.length; i++) {
    // Nama Mapel is now at column index 5
    if (data[i][5] && data[i][5].toString().trim().toUpperCase() === mapel.toUpperCase()) {
      filtered.push(data[i]);
    }
  }
  
  if (filtered.length === 1) return "Tidak ada data untuk mapel: " + mapel;
  return filtered.map(r => r.join(',')).join('\n');
}

/**
 * Download results filtered by Nama Mapel and optionally by Subkelas or Mapel Kelas (base class).
 * Parameters:
 *  - mapel: name of the subject (required)
 *  - subkelas: exact subkelas value as recorded in HASIL (e.g., '7A') (optional)
 *  - mapelKelas: base class value '7'|'8'|'9' to filter by Mapel Kelas (optional)
 * If both subkelas and mapelKelas are provided, subkelas takes precedence.
 */
function downloadHasilPerMapelDanKelas(mapel, subkelas, mapelKelas, npsn) {
  if (!mapel || !mapel.toString().trim()) return 'Parameter mapel diperlukan';
  const sh = ss().getSheetByName('HASIL');
  if (!sh) return 'Data kosong';

  const data = sh.getDataRange().getValues();
  const header = data[0] || [];
  const filtered = [header];
  const targetMapel = mapel.toString().trim().toUpperCase();
  const targetSub = subkelas && subkelas.toString().trim() ? subkelas.toString().trim().toUpperCase() : '';
  const targetBase = mapelKelas && mapelKelas.toString().trim() ? mapelKelas.toString().trim().toUpperCase() : '';
  const targetNpsn = npsn && npsn.toString().trim() ? npsn.toString().trim() : '';

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const namaMapel = (row[5] || '').toString().trim().toUpperCase(); // Nama Mapel
    if (namaMapel !== targetMapel) continue;
    const rowSub = (row[2] || '').toString().trim().toUpperCase(); // Kelas (subkelas)
    const rowBase = (row[4] || '').toString().trim().toUpperCase(); // Mapel Kelas (base)
    const rowNpsn = (row[9] || '').toString().trim(); // NPSN

    if (targetNpsn && rowNpsn !== targetNpsn) continue;
    
    if (targetSub) {
      if (rowSub === targetSub) filtered.push(row);
    } else if (targetBase) {
      if (rowBase === targetBase) filtered.push(row);
    } else {
      filtered.push(row);
    }
  }

  if (filtered.length === 1) return 'Tidak ada data untuk kombinasi yang dipilih';
  return filtered.map(r => r.join(',')).join('\n');
}

/**
 * Download results filtered by kelas (contoh: '7A').
 * Mengembalikan string CSV atau pesan ketika tidak ada data.
 */
function downloadHasilPerKelas(kelas) {
  if (!kelas || !kelas.toString().trim()) return 'Parameter kelas diperlukan';
  const sh = ss().getSheetByName('HASIL');
  if (!sh) return 'Data kosong';

  const data = sh.getDataRange().getValues();
  const filtered = [data[0]]; // header
  const target = kelas.toString().trim().toUpperCase();

  for (let i = 1; i < data.length; i++) {
    const rowKelas = (data[i][2] || '').toString().trim().toUpperCase(); // kolom KELAS berada di index 2
    if (rowKelas === target) filtered.push(data[i]);
  }

  if (filtered.length === 1) return 'Tidak ada data untuk kelas: ' + kelas;
  return filtered.map(r => r.join(',')).join('\n');
}

/**
 * Return list of sub-classes present in sheet HASIL (e.g., 7A,7B,...)
 */
function getSubkelasList() {
  const sh = ss().getSheetByName('HASIL');
  if (!sh || sh.getLastRow() <= 1) return [];
  const data = sh.getDataRange().getValues();
  const set = new Set();
  for (let i = 1; i < data.length; i++) {
    const k = (data[i][2] || '').toString().trim(); // Kelas at index 2
    if (k) set.add(k);
  }
  return Array.from(set).sort();
}

/**
 * Return unique mapel names for a given base class (7/8/9) or for a specific subkelas.
 * If `mapelKelas` provided, collect mapel where Mapel Kelas (col 4) matches.
 * If `subkelas` provided, collect mapel where Kelas (col 2) matches the subkelas.
 */
function getMapelByBaseClass(mapelKelas, subkelas) {
  // Prefer HASIL sheet (if available) as primary source of mapel names
  const set = new Set();
  const targetBase = mapelKelas && mapelKelas.toString().trim() ? mapelKelas.toString().trim().toUpperCase() : '';
  const targetSubRaw = subkelas && subkelas.toString().trim() ? subkelas.toString().trim().toUpperCase() : '';
  // normalize target: if a single digit like '7' treat as base-class filter
  const targetIsBaseDigit = targetSubRaw && targetSubRaw.length === 1 && /^[0-9]$/.test(targetSubRaw);
  const targetSub = targetIsBaseDigit ? '' : targetSubRaw;

  try {
    const sh = ss().getSheetByName('HASIL');
    if (sh && sh.getLastRow() > 1) {
      const data = sh.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rowMapel = (row[5] || '').toString().trim();
        const rowSub = (row[2] || '').toString().trim().toUpperCase();
        const rowBase = (row[4] || '').toString().trim().toUpperCase();
        if (targetSub) {
          if (rowSub === targetSub && rowMapel) set.add(rowMapel);
        } else if (targetBase) {
          if (rowBase === targetBase && rowMapel) set.add(rowMapel);
        } else if (targetIsBaseDigit) {
          // if caller provided base digit (eg '7'), include rows whose subkelas starts with that digit
          if (rowSub && rowSub.charAt(0) === targetSubRaw && rowMapel) set.add(rowMapel);
        }
      }
    }
  } catch (e) {
    // ignore and try fallbacks
  }

  // If no mapel found in HASIL, fallback to known SOAL_* sheets, SOAL_CBT and FORM_LINKS
  if (set.size === 0) {
    try {
      const soalList = listSoalSheets();
      soalList.forEach(obj => {
        const k = (obj.kelas || '').toString().trim().toUpperCase();
        const m = (obj.mapel || '').toString().trim();
        if (!m) return;
        if (targetSub) {
          if (k === targetSub) set.add(m);
        } else if (targetBase) {
          if (k === targetBase) set.add(m);
        } else if (targetIsBaseDigit) {
          if (k && k.charAt(0) === targetSubRaw) set.add(m);
        } else {
          set.add(m);
        }
      });
    } catch (e) { /* ignore */ }

    // also try SOAL_CBT and FORM_LINKS as additional sources
    try {
      const cbt = ss().getSheetByName('SOAL_CBT');
      if (cbt && cbt.getLastRow() > 1) {
        const d = cbt.getDataRange().getValues();
        for (let i = 1; i < d.length; i++) {
          const rowK = (d[i][0] || '').toString().trim().toUpperCase();
          const rowM = (d[i][1] || '').toString().trim();
          if (!rowM) continue;
          if (targetSub) {
            if (rowK === targetSub) set.add(rowM);
          } else if (targetBase) {
            if (rowK === targetBase) set.add(rowM);
          } else if (targetIsBaseDigit) {
            if (rowK && rowK.charAt(0) === targetSubRaw) set.add(rowM);
          } else {
            set.add(rowM);
          }
        }
      }
    } catch (e) { /* ignore */ }

    try {
      const fl = ss().getSheetByName('FORM_LINKS');
      if (fl && fl.getLastRow() > 1) {
        const d = fl.getDataRange().getValues();
        for (let i = 1; i < d.length; i++) {
          const rowK = (d[i][0] || '').toString().trim().toUpperCase();
          const rowM = (d[i][1] || '').toString().trim();
          if (!rowM) continue;
          if (targetSub) {
            if (rowK === targetSub) set.add(rowM);
          } else if (targetBase) {
            if (rowK === targetBase) set.add(rowM);
          } else if (targetIsBaseDigit) {
            if (rowK && rowK.charAt(0) === targetSubRaw) set.add(rowM);
          } else {
            set.add(rowM);
          }
        }
      }
    } catch (e) { /* ignore */ }
  }

  return Array.from(set).sort();
}

/**
 * Return unique mapel names from sheet SOAL_CBT filtered by kelas (base class like '7','8','9').
 * If mapelKelas is empty, returns all mapel listed in SOAL_CBT.
 */
function getMapelFromSoalCBT(mapelKelas) {
  try {
    const sh = ss().getSheetByName('SOAL_CBT');
    if (!sh || sh.getLastRow() <= 1) return [];
    const data = sh.getDataRange().getValues();
    const set = new Set();
    const target = mapelKelas && mapelKelas.toString().trim() ? mapelKelas.toString().trim().toUpperCase() : '';
    for (let i = 1; i < data.length; i++) {
      const rowK = (data[i][0] || '').toString().trim().toUpperCase();
      const rowM = (data[i][1] || '').toString().trim();
      if (!rowM) continue;
      if (target) {
        // match exact kelas or when rowK starts with target (eg '7A' starts with '7')
        if (rowK === target || (rowK && rowK.charAt(0) === target.charAt(0))) set.add(rowM);
      } else {
        set.add(rowM);
      }
    }
    return Array.from(set).sort();
  } catch (e) {
    return [];
  }
}

/**
 * Produce a pivot CSV where columns are Mapel names and rows are students (Nomor,Kelas,Nama).
 * Parameters:
 *  - mapelKelas: base class (7/8/9) to filter by (optional if subkelas provided)
 *  - subkelas: exact subkelas (e.g., '7A') to filter by (optional)
 * Returns CSV string or a message when no data.
 */
function downloadHasilPivot(mapelKelas, subkelas) {
  const sh = ss().getSheetByName('HASIL');
  if (!sh) return 'Data kosong';
  const data = sh.getDataRange().getValues();
  if (!data || data.length <= 1) return 'Data kosong';

  const targetSub = subkelas && subkelas.toString().trim() ? subkelas.toString().trim().toUpperCase() : '';
  const targetBase = mapelKelas && mapelKelas.toString().trim() ? mapelKelas.toString().trim().toUpperCase() : '';
  if (!targetSub && !targetBase) return 'Parameter mapelKelas atau subkelas diperlukan';

  // Determine mapel headers (unique mapel names)
  const mapelSet = new Set();
  // Collect student map: key -> { absen, kelas, nama, latestTs, scores: {mapel: {score, ts}} }
  const students = {};

  const parseTs = (v) => {
    try { const d = new Date(v); const t = d.getTime(); return isNaN(t) ? 0 : t; } catch (e) { return 0; }
  };

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const tsRaw = row[0];
    const ts = parseTs(tsRaw);
    const absen = (row[1] || '').toString().trim();
    const kelasVal = (row[2] || '').toString().trim();
    const nama = (row[3] || '').toString().trim();
    const base = (row[4] || '').toString().trim().toUpperCase();
    const mapel = (row[5] || '').toString().trim();
    const skor = (row[8] || '').toString().trim();

    const rowSub = kelasVal.toUpperCase();

    if (targetSub) {
      if (rowSub !== targetSub) continue;
    } else if (targetBase) {
      if (base !== targetBase) continue;
    }

    if (mapel) mapelSet.add(mapel);

    // student key prefer absen if present, else nama+kelas
    const key = (absen ? ('#' + absen) : '') + '::' + nama + '::' + kelasVal;
    // derive base class if not present in column 4
    const baseDerived = (base && base !== '') ? base : ( (kelasVal && kelasVal.toString().match(/^([789])/)) ? kelasVal.toString().match(/^([789])/)[1] : '' );
    if (!students[key]) students[key] = { absen: absen || '', kelas: kelasVal || '', nama: nama || '', base: baseDerived || '', latestTs: 0, scores: {} };
    // Track latest timestamp for the student (across all mapel)
    if (!students[key].latestTs || students[key].latestTs < ts) students[key].latestTs = ts;
    // For this student+mapel, keep the latest timestamp row
    if (!students[key].scores[mapel] || (students[key].scores[mapel].ts || 0) < ts) {
      students[key].scores[mapel] = { score: skor, ts: ts };
    }
  }

  const mapels = Array.from(mapelSet).sort();
  if (mapels.length === 0) return 'Tidak ada mapel untuk filter yang dipilih';

  // Build header to align with HASIL sheet (include Timestamp first)
  // HASIL header: ['Timestamp','Nomor Absen/Peserta','Kelas','Nama','Mapel Kelas','Nama Mapel','Benar','Salah','Skor']
  // For pivot we include Timestamp, Nomor Absen/Peserta, Kelas, Nama, then one column per Mapel
  const header = ['TIMESTAMP','NO ABSEN/PESERTA','SUB KELAS','NAMA','BASE KELAS'].concat(mapels);
  const rowsOut = [header.map(h => ('"' + (h || '') + '"')).join(',')];

  // Sort students by kelas, then numeric absen if available, else by nama
  const studentKeys = Object.keys(students);
  studentKeys.sort((a,b) => {
    const A = students[a]; const B = students[b];
    const ka = A.kelas.toString().match(/\d+/); const kb = B.kelas.toString().match(/\d+/);
    const na = ka ? parseInt(ka[0],10) : 999; const nb = kb ? parseInt(kb[0],10) : 999;
    if (na !== nb) return na - nb;
    // if both have absen numbers, sort by that
    const aa = A.absen ? parseInt(A.absen.toString().replace(/[^0-9]/g,''),10) : NaN;
    const bb = B.absen ? parseInt(B.absen.toString().replace(/[^0-9]/g,''),10) : NaN;
    if (!isNaN(aa) && !isNaN(bb)) return aa - bb;
    return (A.nama || '').toString().localeCompare((B.nama || '').toString());
  });

  studentKeys.forEach(k => {
    const s = students[k];
    // Format timestamp to readable string (use script timezone)
    const tsVal = s.latestTs ? Utilities.formatDate(new Date(s.latestTs), Session.getScriptTimeZone() || 'GMT+7', 'yyyy-MM-dd HH:mm') : '';
    const cols = [tsVal, s.absen || '', s.kelas || '', s.nama || '', s.base || ''];
    mapels.forEach(m => {
      const val = (s.scores[m] && typeof s.scores[m].score !== 'undefined') ? s.scores[m].score : '';
      cols.push(val);
    });
    // escape commas/quotes
    const esc = cols.map(v => ('"' + (v || '') .toString().replace(/"/g,'""') + '"'));
    rowsOut.push(esc.join(','));
  });

  return rowsOut.join('\n');
}

function downloadFormResponses(kelas, mapel) {
  if (!kelas || !mapel) return 'Parameter kelas/mapel diperlukan';
  const sheetName = 'FORM_RESPONSE_' + kelas.toString().toUpperCase().trim() + '_' + mapel.toString().toUpperCase().trim();
  const sh = ss().getSheetByName(sheetName);
  if (!sh) return 'Tidak ada sheet respons untuk: ' + sheetName;
  const data = sh.getDataRange().getValues();
  if (!data || data.length <= 1) return 'Tidak ada respons di sheet: ' + sheetName;

  const headers = data[0].map(h => (h || '').toString().trim().toUpperCase());
  const findIdx = names => {
    const up = names.map(n => n.toString().toUpperCase());
    for (let i = 0; i < headers.length; i++) {
      for (let n of up) {
        if (headers[i].indexOf(n) !== -1) return i;
      }
    }
    return -1;
  };

  const idxTimestamp = findIdx(['TIMESTAMP', 'WAKTU', 'TANGGAL']);
  const idxNama = findIdx(['NAMA', 'NAME']);
  const idxKelas = findIdx(['KELAS', 'CLASS']);
  const idxAbsen = findIdx(['ABSEN', 'NO', 'NOMOR', 'NOMOR ABSEN', 'NOMOR_ABSEN']);
  const idxSkor = findIdx(['SKOR', 'SCORE']);

  const escape = v => {
    if (v === null || v === undefined) return '';
    const s = v.toString();
    if (s.indexOf(',') !== -1 || s.indexOf('"') !== -1 || s.indexOf('\n') !== -1) return '"' + s.replace(/"/g, '""') + '"';
    return s;
  };

  const outHeader = ['Tanggal', 'Nomor', 'Nama', 'Kelas', 'Mapel', 'Skor'];
  const rows = [outHeader.join(',')];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const tanggal = idxTimestamp !== -1 ? row[idxTimestamp] : row[0];
    let nomor = '';
    if (idxAbsen !== -1) nomor = row[idxAbsen]; else nomor = i; // fallback to serial
    const nama = idxNama !== -1 ? row[idxNama] : '';
    const kelasVal = idxKelas !== -1 ? row[idxKelas] : '';
    const skor = idxSkor !== -1 ? row[idxSkor] : '';
    const mapelVal = mapel;
    rows.push([escape(tanggal), escape(nomor), escape(nama), escape(kelasVal), escape(mapelVal), escape(skor)].join(','));
  }

  return rows.join('\n');
}

function createFormResponseSheet(kelas, mapel) {
  const sheetName = 'FORM_RESPONSE_' + kelas.toUpperCase().trim() + '_' + mapel.toUpperCase().trim();
  let sh = ss().getSheetByName(sheetName);
  
  if (!sh) {
    sh = ss().insertSheet(sheetName);
    sh.appendRow(['Timestamp', 'Nama', 'Kelas', 'Absen', 'Pertanyaan', 'Jawaban', 'Skor']);
    return 'Sheet response berhasil dibuat: ' + sheetName + '. Hubungkan Google Form Anda ke sheet ini.';
  }
  
  return 'Sheet response sudah ada: ' + sheetName;
}

function setAppTitle(title) {
  if (!title || !title.trim()) throw new Error('Judul tidak boleh kosong');
  setSetting('APP_TITLE', title.trim());
  return 'OK';
}

function getAppTitle() {
  var t = getSetting('APP_TITLE');
  return t && t.trim() ? t : 'PAI Smart Assessment 2026';
}

function setAdminPassword(password) {
  if (!password || !password.trim()) throw new Error('Password tidak boleh kosong');
  setSetting('ADMIN_PASSWORD', password.trim());
  return 'OK';
}

function getAdminPassword() {
  var p = getSetting('ADMIN_PASSWORD');
  return p && p.trim() ? p : 'Stuka';
}

function doGet() {
  var appTitle = getAppTitle();
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle(appTitle)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
