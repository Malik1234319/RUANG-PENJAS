/*******************************************************
 * RUANG PENJAS
 * Sistem Administrasi Guru PJOK
 * Google Apps Script + Google Sheets
 *******************************************************/

const APP = {
  NAME: 'RUANG PENJAS',
  YEAR: '2026/2027',
  KKM: 75,
  CLASSES: ['1A', '1B', '1C', '5A', '5B', '6A', '6B'],
  SESSION_SECONDS: 21600
};

const SHEETS = {
  USERS: 'USERS',
  SISWA: 'SISWA',
  KELAS: 'KELAS',
  SETTINGS: 'SETTINGS',
  CP: 'CP_TP_KKTP',
  PROTA: 'PROTA',
  PROMES: 'PROMES',
  JURNAL: 'JURNAL',
  NILAI: 'NILAI',
  REMEDIAL: 'REMEDIAL_PENGAYAAN',
  LAPORAN: 'LAPORAN_NILAI'
};

const HEADERS = {
  USERS: [
    'ID',
    'Username',
    'Password',
    'Nama',
    'Role',
    'Status'
  ],

  SISWA: [
    'ID',
    'Nama',
    'Kelas',
    'NISN',
    'JenisKelamin',
    'NamaOrangTua',
    'NoWAOrangTua',
    'Status'
  ],

  KELAS: [
    'Kelas',
    'Status'
  ],

  SETTINGS: [
    'Kunci',
    'Nilai'
  ],

  CP: [
    'ID',
    'Kelas',
    'Fase',
    'MataPelajaran',
    'CP',
    'TP',
    'KKTP',
    'Kriteria',
    'Semester',
    'TahunAjaran',
    'DibuatOleh'
  ],

  PROTA: [
    'ID',
    'Kelas',
    'Semester',
    'Bulan',
    'Materi',
    'JP',
    'Keterangan',
    'TahunAjaran',
    'DibuatOleh'
  ],

  PROMES: [
    'ID',
    'Kelas',
    'Semester',
    'Minggu',
    'Materi',
    'JP',
    'Keterangan',
    'TahunAjaran',
    'DibuatOleh'
  ],

  JURNAL: [
    'ID',
    'Tanggal',
    'Kelas',
    'Materi',
    'Tujuan',
    'Kegiatan',
    'Kehadiran',
    'Refleksi',
    'TindakLanjut',
    'Guru'
  ],

  NILAI: [
    'ID',
    'Tanggal',
    'Kelas',
    'Nama',
    'NISN',
    'Asesmen',
    'Jenis',
    'Nilai',
    'Keterangan',
    'Guru'
  ],

  REMEDIAL: [
    'ID',
    'Tanggal',
    'Kelas',
    'Nama',
    'NISN',
    'Jenis',
    'Materi',
    'NilaiAwal',
    'NilaiAkhir',
    'TindakLanjut',
    'Guru'
  ],

  LAPORAN: [
    'Kelas',
    'Nama',
    'NISN',
    'JumlahAsesmen',
    'RataRata',
    'Status',
    'TanggalCetak'
  ]
};


/* =====================================================
   WEB APP
===================================================== */

function doGet() {
  initializeSheets_();

  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle(APP.NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


function include(filename) {
  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();
}


/* =====================================================
   SPREADSHEET
===================================================== */

function getSS_() {

  const props = PropertiesService.getScriptProperties();
  const storedId = props.getProperty('SPREADSHEET_ID');

  if (storedId) {
    return SpreadsheetApp.openById(storedId);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error(
      'Spreadsheet belum terhubung. Jalankan script dari Google Spreadsheet yang digunakan.'
    );
  }

  props.setProperty('SPREADSHEET_ID', ss.getId());

  return ss;
}


function setSpreadsheetId(id) {

  if (!id) {
    throw new Error('ID Spreadsheet kosong.');
  }

  PropertiesService
    .getScriptProperties()
    .setProperty('SPREADSHEET_ID', id);

  return 'Spreadsheet berhasil disimpan.';
}


/* =====================================================
   INITIALIZE
===================================================== */

function initializeSheets_() {

  const ss = getSS_();

  Object.keys(SHEETS).forEach(function(key) {

    const name = SHEETS[key];

    let sh = ss.getSheetByName(name);

    if (!sh) {
      sh = ss.insertSheet(name);
    }

    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS[key]);
      sh.setFrozenRows(1);
    }
  });

  seedSettings_();
  seedClasses_();
  seedUsers_();
}


function setupApp() {

  initializeSheets_();

  return {
    success: true,
    message: 'RUANG PENJAS berhasil disiapkan.'
  };
}


function seedSettings_() {

  const sh = getSS_().getSheetByName(SHEETS.SETTINGS);

  if (sh.getLastRow() > 1) return;

  sh.appendRow(['TahunAjaran', APP.YEAR]);
  sh.appendRow(['KKM', APP.KKM]);
  sh.appendRow(['MataPelajaran', 'PJOK']);
}


function seedClasses_() {

  const sh = getSS_().getSheetByName(SHEETS.KELAS);

  if (sh.getLastRow() > 1) return;

  APP.CLASSES.forEach(function(kelas) {
    sh.appendRow([kelas, 'Aktif']);
  });
}


function seedUsers_() {

  const sh = getSS_().getSheetByName(SHEETS.USERS);

  if (sh.getLastRow() > 1) return;

  sh.appendRow([
    Utilities.getUuid(),
    'admin',
    'admin123',
    'Administrator',
    'Admin',
    'Aktif'
  ]);

  sh.appendRow([
    Utilities.getUuid(),
    'guru',
    'guru123',
    'Guru PJOK',
    'Guru',
    'Aktif'
  ]);
}


/* =====================================================
   UTILITIES
===================================================== */

function normalize_(value) {

  return String(value || '')
    .trim()
    .replace(/\s+/g, ' ')
    .toUpperCase();
}


function safeNumber_(value) {

  const n = Number(value);

  return isNaN(n) ? 0 : n;
}


function formatDate_(date) {

  if (!date) return '';

  const d = new Date(date);

  if (isNaN(d.getTime())) {
    return String(date);
  }

  return Utilities.formatDate(
    d,
    Session.getScriptTimeZone() || 'Asia/Jakarta',
    'yyyy-MM-dd'
  );
}


function rowsToObjects_(sheetName) {

  const sh = getSS_().getSheetByName(sheetName);

  if (!sh || sh.getLastRow() < 2) {
    return [];
  }

  const values = sh.getDataRange().getValues();

  const headers = values[0];

  return values.slice(1).map(function(row) {

    const obj = {};

    headers.forEach(function(header, i) {

      let value = row[i];

      if (value instanceof Date) {
        value = formatDate_(value);
      }

      obj[header] = value;
    });

    return obj;
  });
}


function appendObject_(sheetName, headers, obj) {

  const sh = getSS_().getSheetByName(sheetName);

  const row = headers.map(function(header) {
    return obj[header] !== undefined
      ? obj[header]
      : '';
  });

  sh.appendRow(row);
}


/* =====================================================
   SESSION
===================================================== */

function createSession_(user) {

  const token = Utilities.getUuid();

  CacheService
    .getScriptCache()
    .put(
      'SESSION_' + token,
      JSON.stringify(user),
      APP.SESSION_SECONDS
    );

  return token;
}


function getSession_(token) {

  if (!token) {
    throw new Error('Sesi tidak ditemukan.');
  }

  const raw = CacheService
    .getScriptCache()
    .get('SESSION_' + token);

  if (!raw) {
    throw new Error(
      'Sesi telah berakhir. Silakan login kembali.'
    );
  }

  return JSON.parse(raw);
}


function requireAuth_(token, roles) {

  const user = getSession_(token);

  if (roles && roles.indexOf(user.role) === -1) {
    throw new Error('Anda tidak memiliki akses.');
  }

  return user;
}


function logout(token) {

  if (token) {
    CacheService
      .getScriptCache()
      .remove('SESSION_' + token);
  }

  return true;
}


/* =====================================================
   LOGIN GURU / ADMIN
===================================================== */

function loginTeacher(username, password) {

  initializeSheets_();

  const rows = rowsToObjects_(SHEETS.USERS);

  const user = rows.find(function(row) {

    return normalize_(row.Username).toLowerCase() ===
      normalize_(username).toLowerCase()
      &&
      String(row.Password) === String(password)
      &&
      normalize_(row.Status) === 'AKTIF';
  });

  if (!user) {
    throw new Error(
      'Username atau password Guru/Admin salah.'
    );
  }

  const sessionUser = {
    id: user.ID,
    username: user.Username,
    nama: user.Nama,
    role: user.Role
  };

  const token = createSession_(sessionUser);

  return {
    success: true,
    token: token,
    user: sessionUser
  };
}


/* =====================================================
   LOGIN SISWA
===================================================== */

function loginStudent(nama, kelas) {

  initializeSheets_();

  const rows = rowsToObjects_(SHEETS.SISWA);

  const targetNama = normalize_(nama);
  const targetKelas = normalize_(kelas);

  const siswa = rows.find(function(row) {

    return normalize_(row.Nama) === targetNama
      &&
      normalize_(row.Kelas) === targetKelas
      &&
      normalize_(row.Status || 'Aktif') === 'AKTIF';
  });

  if (!siswa) {
    throw new Error(
      'Data siswa tidak ditemukan. Pastikan Nama dan Kelas sesuai.'
    );
  }

  const sessionUser = {
    id: siswa.ID,
    nama: siswa.Nama,
    role: 'Siswa',
    kelas: siswa.Kelas,
    nisn: siswa.NISN || '',
    namaOrangTua: siswa.NamaOrangTua || '',
    noWAOrangTua: siswa.NoWAOrangTua || ''
  };

  const token = createSession_(sessionUser);

  return {
    success: true,
    token: token,
    user: sessionUser
  };
}


/* =====================================================
   APP DATA
===================================================== */

function getAppData(token) {

  const user = requireAuth_(token);

  return {
    user: user,
    classes: getClasses_(),
    settings: getSettings_()
  };
}


function getClasses_() {

  const rows = rowsToObjects_(SHEETS.KELAS);

  const classes = rows
    .filter(function(row) {
      return normalize_(row.Status) === 'AKTIF';
    })
    .map(function(row) {
      return String(row.Kelas).trim();
    });

  return classes.length ? classes : APP.CLASSES;
}


function getSettings_() {

  const rows = rowsToObjects_(SHEETS.SETTINGS);

  const obj = {
    TahunAjaran: APP.YEAR,
    KKM: APP.KKM,
    MataPelajaran: 'PJOK'
  };

  rows.forEach(function(row) {

    if (row.Kunci) {
      obj[row.Kunci] = row.Nilai;
    }
  });

  obj.KKM = safeNumber_(obj.KKM) || APP.KKM;

  return obj;
}


/* =====================================================
   DATA SISWA
===================================================== */

function getStudents(token, kelas) {

  const user = requireAuth_(token);

  let rows = rowsToObjects_(SHEETS.SISWA);

  if (user.role === 'Siswa') {

    rows = rows.filter(function(row) {

      return normalize_(row.Nama) === normalize_(user.nama)
        &&
        normalize_(row.Kelas) === normalize_(user.kelas);
    });

  } else if (kelas) {

    rows = rows.filter(function(row) {

      return normalize_(row.Kelas) === normalize_(kelas);
    });
  }

  return rows;
}


/* =====================================================
   SIMPAN SISWA
===================================================== */

function saveStudent(token, data) {

  const user = requireAuth_(token, ['Admin', 'Guru']);

  if (!data.Nama || !data.Kelas) {
    throw new Error('Nama dan kelas wajib diisi.');
  }

  const rows = rowsToObjects_(SHEETS.SISWA);

  const duplicate = rows.some(function(row) {

    return normalize_(row.Nama) === normalize_(data.Nama)
      &&
      normalize_(row.Kelas) === normalize_(data.Kelas);
  });

  if (duplicate) {
    throw new Error(
      'Siswa dengan nama dan kelas tersebut sudah ada.'
    );
  }

  appendObject_(
    SHEETS.SISWA,
    HEADERS.SISWA,
    {
      ID: Utilities.getUuid(),
      Nama: data.Nama,
      Kelas: data.Kelas,
      NISN: data.NISN || '',
      JenisKelamin: data.JenisKelamin || '',
      NamaOrangTua: data.NamaOrangTua || '',
      NoWAOrangTua: normalizeWhatsApp_(data.NoWAOrangTua || ''),
      Status: data.Status || 'Aktif'
    }
  );

  return {
    success: true,
    message: 'Data siswa berhasil disimpan.',
    oleh: user.nama
  };
}


/* =====================================================
   IMPOR SISWA - COPY PASTE
===================================================== */

function importStudents(token, textData) {

  requireAuth_(token, ['Admin', 'Guru']);

  if (!textData) {
    throw new Error('Data impor kosong.');
  }

  const lines = String(textData)
    .replace(/\r/g, '')
    .split('\n')
    .filter(function(line) {
      return line.trim() !== '';
    });

  if (lines.length < 2) {
    throw new Error(
      'Data minimal terdiri dari header dan satu baris siswa.'
    );
  }

  const delimiter = lines[0].indexOf('\t') >= 0
    ? '\t'
    : ',';

  const headers = lines[0]
    .split(delimiter)
    .map(function(h) {
      return normalize_(h)
        .replace(/_/g, '')
        .replace(/\./g, '');
    });

  const indexOf = function(possibleNames) {

    for (let i = 0; i < possibleNames.length; i++) {

      const idx = headers.indexOf(
        normalize_(possibleNames[i])
          .replace(/_/g, '')
          .replace(/\./g, '')
      );

      if (idx >= 0) return idx;
    }

    return -1;
  };

  const idxNama = indexOf([
    'Nama',
    'Nama Siswa',
    'NamaSiswa'
  ]);

  const idxKelas = indexOf([
    'Kelas'
  ]);

  const idxNISN = indexOf([
    'NISN'
  ]);

  const idxJK = indexOf([
    'JK',
    'Jenis Kelamin',
    'JenisKelamin'
  ]);

  const idxOrtu = indexOf([
    'Nama Orang Tua',
    'NamaOrangTua',
    'Orang Tua',
    'Wali'
  ]);

  const idxWA = indexOf([
    'No WA',
    'No. WA',
    'Nomor WA',
    'WhatsApp',
    'NoWAOrangTua',
    'No WhatsApp'
  ]);

  if (idxNama === -1 || idxKelas === -1) {

    throw new Error(
      'Kolom Nama dan Kelas wajib tersedia.'
    );
  }

  const existing = rowsToObjects_(SHEETS.SISWA);

  const existingKeys = {};

  existing.forEach(function(row) {

    const key =
      normalize_(row.Nama) +
      '|' +
      normalize_(row.Kelas);

    existingKeys[key] = true;
  });

  const result = {
    total: 0,
    berhasil: 0,
    duplikat: 0,
    gagal: 0,
    detail: []
  };

  const rowsToAdd = [];

  lines.slice(1).forEach(function(line, lineIndex) {

    result.total++;

    const cols = line.split(delimiter);

    const nama = (cols[idxNama] || '').trim();
    const kelas = (cols[idxKelas] || '').trim();

    if (!nama || !kelas) {

      result.gagal++;

      result.detail.push(
        'Baris ' +
        (lineIndex + 2) +
        ': Nama/Kelas kosong.'
      );

      return;
    }

    const key =
      normalize_(nama) +
      '|' +
      normalize_(kelas);

    if (existingKeys[key]) {

      result.duplikat++;

      return;
    }

    const row = {
      ID: Utilities.getUuid(),
      Nama: nama,
      Kelas: kelas,
      NISN: idxNISN >= 0
        ? (cols[idxNISN] || '').trim()
        : '',
      JenisKelamin: idxJK >= 0
        ? (cols[idxJK] || '').trim()
        : '',
      NamaOrangTua: idxOrtu >= 0
        ? (cols[idxOrtu] || '').trim()
        : '',
      NoWAOrangTua: idxWA >= 0
        ? normalizeWhatsApp_(cols[idxWA] || '')
        : '',
      Status: 'Aktif'
    };

    rowsToAdd.push(row);

    existingKeys[key] = true;

    result.berhasil++;
  });

  if (rowsToAdd.length) {

    const sh = getSS_().getSheetByName(SHEETS.SISWA);

    const values = rowsToAdd.map(function(obj) {

      return HEADERS.SISWA.map(function(header) {

        return obj[header] !== undefined
          ? obj[header]
          : '';
      });
    });

    sh
      .getRange(
        sh.getLastRow() + 1,
        1,
        values.length,
        HEADERS.SISWA.length
      )
      .setValues(values);
  }

  return result;
}


/* =====================================================
   NORMALIZE WHATSAPP
===================================================== */

function normalizeWhatsApp_(number) {

  let n = String(number || '')
    .replace(/[^\d+]/g, '');

  if (!n) return '';

  if (n.indexOf('+62') === 0) {
    n = '62' + n.substring(3);
  }

  if (n.indexOf('62') === 0) {
    return n;
  }

  if (n.indexOf('0') === 0) {
    return '62' + n.substring(1);
  }

  return n;
}


/* =====================================================
   WHATSAPP LINK
===================================================== */

function getWhatsAppLink(token, studentId, message) {

  requireAuth_(token, ['Admin', 'Guru']);

  const rows = rowsToObjects_(SHEETS.SISWA);

  const siswa = rows.find(function(row) {
    return String(row.ID) === String(studentId);
  });

  if (!siswa) {
    throw new Error('Data siswa tidak ditemukan.');
  }

  const phone = normalizeWhatsApp_(siswa.NoWAOrangTua);

  if (!phone) {
    throw new Error(
      'Nomor WhatsApp orang tua belum tersedia.'
    );
  }

  const defaultMessage =
    'Assalamu\'alaikum Bapak/Ibu ' +
    (siswa.NamaOrangTua || 'Orang Tua/Wali') +
    '.\n\n' +
    'Kami dari Guru PJOK SD Negeri 206 Palembang ' +
    'menyampaikan informasi terkait perkembangan ' +
    'pembelajaran PJOK ' +
    siswa.Nama +
    ' kelas ' +
    siswa.Kelas +
    '.\n\n' +
    'Terima kasih atas perhatian dan kerja samanya.\n\n' +
    'Wassalamu\'alaikum.';

  const finalMessage = message || defaultMessage;

  return {
    phone: phone,
    url:
      'https://wa.me/' +
      phone +
      '?text=' +
      encodeURIComponent(finalMessage)
  };
}


/* =====================================================
   GENERIC SAVE
===================================================== */

function saveCP(token, data) {

  const user = requireAuth_(token, ['Admin', 'Guru']);

  appendObject_(
    SHEETS.CP,
    HEADERS.CP,
    {
      ID: Utilities.getUuid(),
      Kelas: data.Kelas,
      Fase: data.Fase || '',
      MataPelajaran: 'PJOK',
      CP: data.CP || '',
      TP: data.TP || '',
      KKTP: data.KKTP || '',
      Kriteria: data.Kriteria || '',
      Semester: data.Semester || '',
      TahunAjaran: getSettings_().TahunAjaran,
      DibuatOleh: user.nama
    }
  );

  return true;
}


function saveProta(token, data) {

  const user = requireAuth_(token, ['Admin', 'Guru']);

  appendObject_(
    SHEETS.PROTA,
    HEADERS.PROTA,
    {
      ID: Utilities.getUuid(),
      Kelas: data.Kelas,
      Semester: data.Semester || '',
      Bulan: data.Bulan || '',
      Materi: data.Materi || '',
      JP: data.JP || '',
      Keterangan: data.Keterangan || '',
      TahunAjaran: getSettings_().TahunAjaran,
      DibuatOleh: user.nama
    }
  );

  return true;
}


function savePromes(token, data) {

  const user = requireAuth_(token, ['Admin', 'Guru']);

  appendObject_(
    SHEETS.PROMES,
    HEADERS.PROMES,
    {
      ID: Utilities.getUuid(),
      Kelas: data.Kelas,
      Semester: data.Semester || '',
      Minggu: data.Minggu || '',
      Materi: data.Materi || '',
      JP: data.JP || '',
      Keterangan: data.Keterangan || '',
      TahunAjaran: getSettings_().TahunAjaran,
      DibuatOleh: user.nama
    }
  );

  return true;
}


function saveJournal(token, data) {

  const user = requireAuth_(token, ['Admin', 'Guru']);

  appendObject_(
    SHEETS.JURNAL,
    HEADERS.JURNAL,
    {
      ID: Utilities.getUuid(),
      Tanggal: data.Tanggal || new Date(),
      Kelas: data.Kelas,
      Materi: data.Materi || '',
      Tujuan: data.Tujuan || '',
      Kegiatan: data.Kegiatan || '',
      Kehadiran: data.Kehadiran || '',
      Refleksi: data.Refleksi || '',
      TindakLanjut: data.TindakLanjut || '',
      Guru: user.nama
    }
  );

  return true;
}


function saveAssessment(token, data) {

  const user = requireAuth_(token, ['Admin', 'Guru']);

  const nilai = safeNumber_(data.Nilai);

  if (nilai < 0 || nilai > 100) {
    throw new Error('Nilai harus antara 0 sampai 100.');
  }

  appendObject_(
    SHEETS.NILAI,
    HEADERS.NILAI,
    {
      ID: Utilities.getUuid(),
      Tanggal: data.Tanggal || new Date(),
      Kelas: data.Kelas,
      Nama: data.Nama,
      NISN: data.NISN || '',
      Asesmen: data.Asesmen || '',
      Jenis: data.Jenis || '',
      Nilai: nilai,
      Keterangan: data.Keterangan || '',
      Guru: user.nama
    }
  );

  return true;
}


function saveRemedial(token, data) {

  const user = requireAuth_(token, ['Admin', 'Guru']);

  appendObject_(
    SHEETS.REMEDIAL,
    HEADERS.REMEDIAL,
    {
      ID: Utilities.getUuid(),
      Tanggal: data.Tanggal || new Date(),
      Kelas: data.Kelas,
      Nama: data.Nama,
      NISN: data.NISN || '',
      Jenis: data.Jenis || '',
      Materi: data.Materi || '',
      NilaiAwal: safeNumber_(data.NilaiAwal),
      NilaiAkhir: safeNumber_(data.NilaiAkhir),
      TindakLanjut: data.TindakLanjut || '',
      Guru: user.nama
    }
  );

  return true;
}


/* =====================================================
   GET DATA MODULE
===================================================== */

function getModuleData(token, module, kelas) {

  const user = requireAuth_(token);

  const map = {
    cp: SHEETS.CP,
    prota: SHEETS.PROTA,
    promes: SHEETS.PROMES,
    jurnal: SHEETS.JURNAL,
    nilai: SHEETS.NILAI,
    remedial: SHEETS.REMEDIAL
  };

  if (!map[module]) {
    throw new Error('Modul tidak ditemukan.');
  }

  let rows = rowsToObjects_(map[module]);

  if (user.role === 'Siswa') {

    if (module === 'nilai') {

      rows = rows.filter(function(row) {
        return normalize_(row.Nama) === normalize_(user.nama)
          &&
          normalize_(row.Kelas) === normalize_(user.kelas);
      });

    } else if (module === 'remedial') {

      rows = rows.filter(function(row) {
        return normalize_(row.Nama) === normalize_(user.nama)
          &&
          normalize_(row.Kelas) === normalize_(user.kelas);
      });

    } else {

      rows = rows.filter(function(row) {
        return normalize_(row.Kelas) === normalize_(user.kelas);
      });
    }

  } else if (kelas) {

    rows = rows.filter(function(row) {

      return normalize_(row.Kelas) === normalize_(kelas);
    });
  }

  return rows;
}


/* =====================================================
   DASHBOARD
===================================================== */

function getDashboard(token, kelas) {

  const user = requireAuth_(token);

  if (user.role === 'Siswa') {
    kelas = user.kelas;
  }

  const students = rowsToObjects_(SHEETS.SISWA)
    .filter(function(row) {

      return normalize_(row.Kelas) === normalize_(kelas);
    });

  const nilai = rowsToObjects_(SHEETS.NILAI)
    .filter(function(row) {

      return normalize_(row.Kelas) === normalize_(kelas);
    });

  const jurnal = rowsToObjects_(SHEETS.JURNAL)
    .filter(function(row) {

      return normalize_(row.Kelas) === normalize_(kelas);
    });

  const cp = rowsToObjects_(SHEETS.CP)
    .filter(function(row) {

      return normalize_(row.Kelas) === normalize_(kelas);
    });

  let totalNilai = 0;

  nilai.forEach(function(row) {
    totalNilai += safeNumber_(row.Nilai);
  });

  const rataRata = nilai.length
    ? totalNilai / nilai.length
    : 0;

  const studentMap = {};

  nilai.forEach(function(row) {

    const key = normalize_(row.Nama);

    if (!studentMap[key]) {
      studentMap[key] = [];
    }

    studentMap[key].push(
      safeNumber_(row.Nilai)
    );
  });

  let tuntas = 0;
  let remedial = 0;

  Object.keys(studentMap).forEach(function(key) {

    const arr = studentMap[key];

    const avg =
      arr.reduce(function(a, b) {
        return a + b;
      }, 0) / arr.length;

    if (avg >= Number(getSettings_().KKM)) {
      tuntas++;
    } else {
      remedial++;
    }
  });

  return {
    kelas: kelas,
    jumlahSiswa: students.length,
    jumlahNilai: nilai.length,
    rataRata: Number(rataRata.toFixed(2)),
    tuntas: tuntas,
    remedial: remedial,
    jurnal: jurnal.length,
    cp: cp.length
  };
}


/* =====================================================
   LAPORAN NILAI
===================================================== */

function getReport(token, kelas) {

  const user = requireAuth_(token);

  if (user.role === 'Siswa') {
    kelas = user.kelas;
  }

  const students = rowsToObjects_(SHEETS.SISWA)
    .filter(function(row) {

      return normalize_(row.Kelas) === normalize_(kelas);
    });

  const nilai = rowsToObjects_(SHEETS.NILAI)
    .filter(function(row) {

      return normalize_(row.Kelas) === normalize_(kelas);
    });

  const kkm = Number(getSettings_().KKM) || APP.KKM;

  const report = [];

  students.forEach(function(student) {

    const values = nilai.filter(function(row) {

      return normalize_(row.Nama) ===
        normalize_(student.Nama);
    });

    const total = values.reduce(function(sum, row) {

      return sum + safeNumber_(row.Nilai);

    }, 0);

    const avg = values.length
      ? total / values.length
      : 0;

    let status = 'Belum Dinilai';

    if (values.length) {
      status = avg >= kkm
        ? 'Tuntas'
        : 'Remedial';
    }

    report.push({
      kelas: student.Kelas,
      nama: student.Nama,
      nisn: student.NISN || '',
      jumlahAsesmen: values.length,
      rataRata: Number(avg.toFixed(2)),
      status: status,
      namaOrangTua: student.NamaOrangTua || '',
      noWAOrangTua: student.NoWAOrangTua || '',
      id: student.ID
    });
  });

  if (user.role === 'Siswa') {

    return report.filter(function(row) {

      return normalize_(row.nama) === normalize_(user.nama);
    });
  }

  return report;
}


/* =====================================================
   GENERATE SHEET LAPORAN
===================================================== */

function generateReportSheet(token, kelas) {

  requireAuth_(token, ['Admin', 'Guru']);

  const report = getReport(token, kelas);

  const sh = getSS_().getSheetByName(SHEETS.LAPORAN);

  if (sh.getLastRow() > 1) {

    sh.getRange(
      2,
      1,
      sh.getLastRow() - 1,
      sh.getLastColumn()
    ).clearContent();
  }

  if (!report.length) {
    return {
      success: false,
      message: 'Belum ada data siswa.'
    };
  }

  const now = formatDate_(new Date());

  const values = report.map(function(row) {

    return [
      row.kelas,
      row.nama,
      row.nisn,
      row.jumlahAsesmen,
      row.rataRata,
      row.status,
      now
    ];
  });

  sh
    .getRange(
      2,
      1,
      values.length,
      HEADERS.LAPORAN.length
    )
    .setValues(values);

  sh.autoResizeColumns(
    1,
    HEADERS.LAPORAN.length
  );

  return {
    success: true,
    message:
      'Laporan kelas ' +
      kelas +
      ' berhasil dibuat.'
  };
}


/* =====================================================
   DELETE DATA
===================================================== */

function deleteRecord(token, module, id) {

  requireAuth_(token, ['Admin', 'Guru']);

  const map = {
    siswa: SHEETS.SISWA,
    cp: SHEETS.CP,
    prota: SHEETS.PROTA,
    promes: SHEETS.PROMES,
    jurnal: SHEETS.JURNAL,
    nilai: SHEETS.NILAI,
    remedial: SHEETS.REMEDIAL
  };

  const sheetName = map[module];

  if (!sheetName) {
    throw new Error('Modul tidak valid.');
  }

  const sh = getSS_().getSheetByName(sheetName);

  const data = sh.getDataRange().getValues();

  if (data.length < 2) {
    throw new Error('Data tidak ditemukan.');
  }

  const headers = data[0];

  const idIndex = headers.indexOf('ID');

  for (let i = 1; i < data.length; i++) {

    if (String(data[i][idIndex]) === String(id)) {

      sh.deleteRow(i + 1);

      return true;
    }
  }

  throw new Error('Data tidak ditemukan.');
}


/* =====================================================
   TEMPLATE IMPOR
===================================================== */

function getImportTemplate() {

  return [
    'Nama Siswa\tKelas\tNISN\tJenis Kelamin\tNama Orang Tua\tNo WA\tStatus',
    'Ahmad Fajar\t6A\t1234567890\tL\tBapak Ahmad\t08123456789\tAktif',
    'Siti Aisyah\t6A\t1234567891\tP\tIbu Siti\t08129876543\tAktif'
  ].join('\n');
}
