/**
 * SISTEM INFORMASI LAPORAN PENGGUNAAN BBM INTENSITAS
 * Google Apps Script + Google Sheets + HTML Service
 *
 * Versi: Final 1.0
 * Unit : KODIM 0807/TULUNGAGUNG
 *
 * Struktur sheet:
 *  - DRIVER
 *  - KENDARAAN
 *  - TUJUAN
 *  - LAPORAN
 *  - USERS (akun ADMIN)
 *
 * Login DRIVER:
 *  Username = Id_Driver
 *  Password = NIP/NRP
 *
 * Login ADMIN:
 *  Username = adminlog
 *  Password = dikonfigurasi pada APP.ADMIN_PASSWORD
 *
 * CATATAN:
 * 1. Jalankan SETUP_SISTEM() satu kali dari editor Apps Script.
 * 2. Buka URL database yang dikembalikan pada log untuk melihat Google Sheet.
 * 3. Deploy sebagai Web App: Execute as Me, Who has access sesuai kebijakan instansi.
 */

const APP = Object.freeze({
  NAME: 'SISTEM INFORMASI LAPORAN PENGGUNAAN BBM INTENSITAS',
  UNIT: 'KODIM 0807/TULUNGAGUNG',
  VERSION: '1.0.0',
  MAX_UPLOAD_BYTES: 5 * 1024 * 1024,
  SESSION_SECONDS: 21600,
  DB_ID_PROP: 'BBM_DB_ID',
  SURAT_FOLDER_PROP: 'BBM_SURAT_FOLDER_ID',
  EXPORT_FOLDER_PROP: 'BBM_EXPORT_FOLDER_ID',
  FIXED_DB_ID: '1_b28MuJ-SlTzbX37MHzPDIpc74iP-dOrWsAYRmQJuW4',
  FIXED_SURAT_FOLDER_ID: '13yhTjsSIUxKngYLR06McPcDWHKmm6yVC',
  ADMIN_USERNAME: 'adminlog',
  ADMIN_PASSWORD: 'Log!stik0807',
  ADMIN_NAME: 'Administrator',
  CLAMP_NEGATIVE_SUPPORT_TO_ZERO: true
});

const SHEET = Object.freeze({
  DRIVER: 'DRIVER',
  KENDARAAN: 'KENDARAAN',
  TUJUAN: 'TUJUAN',
  LAPORAN: 'LAPORAN',
  USERS: 'USERS'
});

const HEADERS = Object.freeze({
  DRIVER: ['Id_Driver', 'Nama_Lengkap', 'Pangkat_Korps', 'NIP_NRP', 'Status'],
  KENDARAAN: ['Id_Kendaraan', 'No_Pol', 'Merk_Type', 'Jumlah_Km_Per_Liter', 'BBM_Rutin', 'Jenis_BBM'],
  TUJUAN: ['Id_Tujuan', 'Nama_Tujuan', 'Jarak_PP'],
  LAPORAN: [
    'Id_Laporan', 'Hari_Tanggal', 'Id_Driver', 'Nama_Driver',
    'Uraian_Kegiatan', 'Id_Kendaraan', 'Kendaraan', 'Id_Tujuan',
    'Tujuan', 'Jarak_PP', 'Km_Per_Liter', 'BBM_Rutin', 'Jenis_BBM',
    'Dukungan_BBM', 'Surat_Tugas_File_Id', 'Surat_Tugas_Nama', 'Created_At'
  ],
  USERS: ['Username', 'Password', 'Role', 'Nama', 'Status', 'Created_At']
});

// =============================
// ENTRY POINT
// =============================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP.NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// =============================
// INITIAL SETUP
// =============================
function SETUP_SISTEM() {
  const props = PropertiesService.getScriptProperties();

  // Gunakan Google Spreadsheet yang sudah ditentukan pengguna.
  const dbId = APP.FIXED_DB_ID;
  props.setProperty(APP.DB_ID_PROP, dbId);
  const ss = SpreadsheetApp.openById(dbId);

  Object.keys(SHEET).forEach(function(key) {
    const sheetName = SHEET[key];
    let sh = ss.getSheetByName(sheetName);
    if (!sh) sh = ss.insertSheet(sheetName);
    const header = HEADERS[key];
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, header.length).setValues([header]);
    } else {
      const existing = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), header.length)).getValues()[0];
      if (existing.slice(0, header.length).join('|') !== header.join('|')) {
        sh.getRange(1, 1, 1, header.length).setValues([header]);
      }
    }
    formatHeader_(sh, header.length);
  });

  // Pastikan sheet USERS tersedia, tetapi JANGAN menulis ke baris/sel yang mungkin
  // sudah memiliki validasi data milik pengguna. Login admin menggunakan akun tetap
  // pada APP dan fallback dari USERS tetap didukung di fungsi login().
  ensureUsersSheetSafe_(ss);

  // Gunakan folder Surat Tugas yang sudah ditentukan pengguna.
  props.setProperty(APP.SURAT_FOLDER_PROP, APP.FIXED_SURAT_FOLDER_ID);
  const suratFolder = DriveApp.getFolderById(APP.FIXED_SURAT_FOLDER_ID);

  // Folder export lokal sistem tetap tersedia sebagai cadangan/arsip bila diperlukan.
  const exportFolder = getOrCreateFolder_(props, APP.EXPORT_FOLDER_PROP, 'BBM_INTENSITAS_EXPORT');

  ss.setActiveSheet(ss.getSheetByName(SHEET.DRIVER));
  SpreadsheetApp.flush();

  Logger.log('Database URL: ' + ss.getUrl());
  Logger.log('Surat Tugas Folder: ' + suratFolder.getUrl());
  Logger.log('Export Folder: ' + exportFolder.getUrl());
  Logger.log('Admin username: ' + APP.ADMIN_USERNAME);
  return {
    databaseId: dbId,
    databaseUrl: ss.getUrl(),
    adminUsername: APP.ADMIN_USERNAME,
    adminPassword: APP.ADMIN_PASSWORD,
    suratFolderUrl: suratFolder.getUrl(),
    exportFolderUrl: exportFolder.getUrl()
  };
}

function ensureCoreSheets_(ss) {
  Object.keys(SHEET).forEach(function(key) {
    const sheetName = SHEET[key];
    let sh = ss.getSheetByName(sheetName);
    if (!sh) sh = ss.insertSheet(sheetName);
    const header = HEADERS[key];
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, header.length).setValues([header]);
      formatHeader_(sh, header.length);
    }
  });
  // Jangan mengubah data/status USERS atau DRIVER saat login.
  // Status master DRIVER dikelola oleh fungsi CRUD khusus.
}

function ensureDriverStatusColumn_(ss) {
  const sh = ss.getSheetByName(SHEET.DRIVER);
  if (!sh) return;

  const lastCol = sh.getLastColumn();
  const headers = lastCol > 0
    ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String)
    : [];

  let statusCol = headers.findIndex(function(h) {
    return h.trim().toLowerCase() === 'status';
  }) + 1;

  if (!statusCol) {
    statusCol = lastCol + 1;
    sh.getRange(1, statusCol).setValue('Status');
    formatHeader_(sh, statusCol);
  }

  const lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    const range = sh.getRange(2, statusCol, lastRow - 1, 1);
    const values = range.getValues();
    let changed = false;

    values.forEach(function(row) {
      const current = String(row[0] == null ? '' : row[0]).trim().toUpperCase();
      if (!current) {
        row[0] = 'Aktif';
        changed = true;
      } else if (current === 'AKTIF') {
        row[0] = 'Aktif';
        changed = true;
      } else if (current === 'TIDAK AKTIF') {
        row[0] = 'Tidak Aktif';
        changed = true;
      }
    });

    if (changed) range.setValues(values);
  }
}

function ensureUsersSheetSafe_(ss) {
  let sh = ss.getSheetByName(SHEET.USERS);
  if (!sh) sh = ss.insertSheet(SHEET.USERS);

  // Hanya membuat header jika sheet benar-benar kosong.
  // Tidak menyentuh data existing/validasi existing milik pengguna.
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.USERS.length).setValues([HEADERS.USERS]);
    formatHeader_(sh, HEADERS.USERS.length);
  }
}

function ensureAdminAccount_(ss) {
  const sh = ss.getSheetByName(SHEET.USERS);
  if (!sh) throw new Error('Sheet USERS tidak ditemukan.');

  const values = sh.getDataRange().getValues();
  const targetUser = APP.ADMIN_USERNAME.toLowerCase();
  let foundRow = null;

  for (let i = 1; i < values.length; i++) {
    const username = String(values[i][0] || '').trim().toLowerCase();
    if (username === targetUser) {
      foundRow = i + 1;
      break;
    }
  }

  const row = [APP.ADMIN_USERNAME, APP.ADMIN_PASSWORD, 'ADMIN', APP.ADMIN_NAME, 'AKTIF'];
  if (foundRow) {
    sh.getRange(foundRow, 1, 1, row.length).setValues([row]);
  } else {
    sh.appendRow(row);
  }
}

function getOrCreateFolder_(props, propName, folderName) {
  const stored = props.getProperty(propName);
  if (stored) {
    try {
      return DriveApp.getFolderById(stored);
    } catch (e) {}
  }
  const folder = DriveApp.createFolder(folderName);
  props.setProperty(propName, folder.getId());
  return folder;
}

function getDb_() {
  const props = PropertiesService.getScriptProperties();
  // Jangan bergantung pada SETUP_SISTEM() untuk proses login.
  // Jika property belum pernah disimpan, gunakan spreadsheet tetap yang sudah ditentukan.
  const id = props.getProperty(APP.DB_ID_PROP) || APP.FIXED_DB_ID;
  props.setProperty(APP.DB_ID_PROP, id);
  try {
    return SpreadsheetApp.openById(id);
  } catch (e) {
    throw new Error('Database Google Spreadsheet tidak dapat dibuka. Pastikan deployment dijalankan sebagai pemilik project dan akun pemilik memiliki akses ke spreadsheet. Detail: ' + e.message);
  }
}

function formatHeader_(sh, width) {
  if (width < 1) return;
  sh.getRange(1, 1, 1, width)
    .setFontWeight('bold')
    .setFontColor('#FFFFFF')
    .setBackground('#0F5A4F')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, width);
}

function ping() {
  const ss = getDb_();
  return {
    ok: true,
    app: APP.NAME,
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    timezone: Session.getScriptTimeZone()
  };
}

// =============================
// SESSION / AUTHENTICATION
// =============================
function login(username, password) {
  username = clean_(username);
  password = String(password == null ? '' : password).trim();

  if (!username || !password) {
    throw new Error('Username dan password wajib diisi.');
  }

  // ==========================================================
  // ADMIN UTAMA SISTEM
  // Tidak bergantung pada sheet USERS.
  // ==========================================================
  if (
    username.toLowerCase() === APP.ADMIN_USERNAME.toLowerCase() &&
    password === APP.ADMIN_PASSWORD
  ) {
    return issueSession_({
      role: 'ADMIN',
      username: APP.ADMIN_USERNAME,
      displayName: APP.ADMIN_NAME,
      idDriver: '',
      pangkat: '',
      nipNrp: ''
    });
  }

  // ==========================================================
  // LOGIN USER DARI SHEET USERS
  // Struktur FINAL:
  // A Username | B Password | C Role | D Nama |
  // E Status   | F Created_At
  // ==========================================================
  const ss = getDb_();
  const usersSheet = ss.getSheetByName(SHEET.USERS);

  if (!usersSheet || usersSheet.getLastRow() < 2) {
    throw new Error('Data akun pada sheet USERS belum tersedia.');
  }

  const users = readObjects_(usersSheet);
  const usernameLower = username.toLowerCase();

  const account = users.find(function(u) {
    const uUsername = String(u.Username || '').trim().toLowerCase();
    const uPassword = String(u.Password || '').trim();
    const uStatus = String(u.Status || '').trim().toUpperCase();

    return (
      uUsername === usernameLower &&
      uPassword === password &&
      uStatus === 'AKTIF'
    );
  });

  if (!account) {
    throw new Error('Username atau password tidak sesuai.');
  }

  const role = String(account.Role || '').trim().toUpperCase();
  const accountUsername = String(account.Username || username).trim();
  const displayName = String(account.Nama || accountUsername).trim();

  // ==========================================================
  // ADMIN DARI USERS
  // ==========================================================
  if (role === 'ADMIN') {
    return issueSession_({
      role: 'ADMIN',
      username: accountUsername,
      displayName: displayName || APP.ADMIN_NAME,
      idDriver: '',
      pangkat: '',
      nipNrp: ''
    });
  }

  // ==========================================================
  // DRIVER / STAF / AJUDAN
  // ==========================================================
  if (role === 'DRIVER' || role === 'STAF' || role === 'AJUDAN') {
    let idDriver = '';
    let namaUser = displayName;
    let pangkat = '';
    let nipNrp = '';

    // Hanya DRIVER yang harus mempunyai pasangan pada Master DRIVER.
    if (role === 'DRIVER') {
      const driverSheet = ss.getSheetByName(SHEET.DRIVER);

      if (!driverSheet || driverSheet.getLastRow() < 2) {
        throw new Error('Master DRIVER belum memiliki data.');
      }

      const drivers = readObjects_(driverSheet);
      const driver = drivers.find(function(d) {
        const nama = String(d.Nama_Lengkap || '').trim().toLowerCase();
        const id = String(d.Id_Driver || '').trim().toLowerCase();

        // Username USERS dapat berupa Id_Driver atau Nama_Lengkap.
        return nama === usernameLower || id === usernameLower;
      });

      if (!driver) {
        throw new Error(
          'Akun DRIVER ditemukan, tetapi username tidak cocok dengan Id_Driver atau Nama_Lengkap pada Master DRIVER.'
        );
      }

      const driverStatus = String(driver.Status || '').trim().toUpperCase();
      if (driverStatus && driverStatus !== 'AKTIF') {
        throw new Error('Data driver tidak aktif. Silakan hubungi Administrator.');
      }

      idDriver = String(driver.Id_Driver || '').trim();
      namaUser = String(driver.Nama_Lengkap || namaUser).trim();
      pangkat = String(driver.Pangkat_Korps || '').trim();
      nipNrp = String(driver.NIP_NRP || '').trim();

      if (!idDriver) {
        throw new Error('Id_Driver pada Master DRIVER belum tersedia.');
      }
    }

    return issueSession_({
      role: role,
      username: accountUsername,
      displayName: namaUser,
      idDriver: idDriver,
      pangkat: pangkat,
      nipNrp: nipNrp
    });
  }

  throw new Error(
    'Role pengguna tidak valid. Gunakan ADMIN, DRIVER, STAF, atau AJUDAN.'
  );
}

function issueSession_(user) {
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('BBM_SESSION_' + token, JSON.stringify(user), APP.SESSION_SECONDS);
  return { token: token, user: user, expiresIn: APP.SESSION_SECONDS };
}

function getSession(token) {
  if (!token) return null;
  const value = CacheService.getScriptCache().get('BBM_SESSION_' + token);
  if (!value) return null;
  try { return JSON.parse(value); } catch (e) { return null; }
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove('BBM_SESSION_' + token);
  return true;
}

function requireAuth_(token, roles) {
  const user = getSession(token);
  if (!user) throw new Error('Sesi login sudah berakhir. Silakan login kembali.');
  if (roles && roles.length && roles.indexOf(user.role) < 0) throw new Error('Anda tidak memiliki hak akses untuk tindakan ini.');
  return user;
}

// =============================
// BOOTSTRAP / DASHBOARD
// =============================
function getBootstrap(token) {
  const user = requireAuth_(token);
  return {
    app: APP,
    user: user,
    formOptions: getFormOptions_(user),
    dashboard: buildDashboard_(user)
  };
}

function getDashboard(token) {
  const user = requireAuth_(token);
  return buildDashboard_(user);
}

function buildDashboard_(user) {
  const ss = getDb_();
  const drivers = readObjects_(ss.getSheetByName(SHEET.DRIVER));
  const vehicles = readObjects_(ss.getSheetByName(SHEET.KENDARAAN));
  const destinations = readObjects_(ss.getSheetByName(SHEET.TUJUAN));
  const reports = getReportObjects_(ss);
  const filtered = user.role === 'DRIVER'
    ? reports.filter(function(r) { return String(r.Id_Driver) === String(user.idDriver); })
    : reports;

  const monthKey = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  const monthReports = filtered.filter(function(r) {
    return formatDateValue_(r.Hari_Tanggal).slice(0, 7) === monthKey;
  });

  const fuelTotals = { Pertalite: 0, Solar: 0 };
  monthReports.forEach(function(r) {
    const fuel = normalizeFuel_(r.Jenis_BBM);
    fuelTotals[fuel] = round2_(fuelTotals[fuel] + num_(r.Dukungan_BBM));
  });

  const recent = filtered.slice().sort(function(a,b) {
    return String(b.Created_At).localeCompare(String(a.Created_At));
  }).slice(0, 7).map(function(r) { return reportView_(r); });

  return {
    counts: {
      driver: user.role === 'ADMIN' ? drivers.length : 1,
      kendaraan: vehicles.length,
      tujuan: destinations.length,
      laporan: filtered.length
    },
    month: {
      periode: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM yyyy'),
      pertalite: fuelTotals.Pertalite,
      solar: fuelTotals.Solar,
      total: round2_(fuelTotals.Pertalite + fuelTotals.Solar),
      jumlahLaporan: monthReports.length
    },
    recent: recent
  };
}

function getFormOptions(token) {
  const user = requireAuth_(token, ['DRIVER', 'STAF', 'AJUDAN', 'ADMIN']);
  return getFormOptions_(user);
}

function getFormOptions_(user) {
  const ss = getDb_();
  const vehicles = readObjects_(ss.getSheetByName(SHEET.KENDARAAN)).map(function(v) {
    return {
      id: String(v.Id_Kendaraan),
      noPol: String(v.No_Pol),
      merkType: String(v.Merk_Type),
      kmPerLiter: num_(v.Jumlah_Km_Per_Liter),
      bbmRutin: num_(v.BBM_Rutin),
      jenisBbm: normalizeFuel_(v.Jenis_BBM)
    };
  });
  const destinations = readObjects_(ss.getSheetByName(SHEET.TUJUAN)).map(function(t) {
    return {
      id: String(t.Id_Tujuan),
      nama: String(t.Nama_Tujuan),
      jarakPP: num_(t.Jarak_PP)
    };
  });
  const drivers = readObjects_(ss.getSheetByName(SHEET.DRIVER)).map(function(d) {
    return {
      id: String(d.Id_Driver),
      nama: String(d.Nama_Lengkap),
      pangkat: String(d.Pangkat_Korps || ''),
      nipNrp: String(d.NIP_NRP || '')
    };
  });
  return { vehicles: vehicles, destinations: destinations, drivers: drivers };
}

// =============================
// REPORT TRANSACTIONS
// =============================
function saveReport(token, payload) {
  const user = requireAuth_(token, ['DRIVER', 'STAF', 'AJUDAN', 'ADMIN']);
  payload = payload || {};

  const tanggal = validateDateInput_(payload.tanggal);
  const uraian = clean_(payload.uraian);
  if (!uraian) throw new Error('Uraian kegiatan wajib diisi.');

  const manualVehicle = !!payload.manualVehicle;
  const manualDestination = !!payload.manualDestination;
  const kendaraanId = clean_(payload.kendaraanId);
  const tujuanId = clean_(payload.tujuanId);

  if (!manualVehicle && !kendaraanId) throw new Error('Kendaraan wajib dipilih.');
  if (!manualDestination && !tujuanId) throw new Error('Tujuan wajib dipilih.');

  const ss = getDb_();

  // DRIVER selalu menggunakan dirinya sendiri. Admin/Staf/Ajudan memilih driver.
  const driverId = user.role === 'DRIVER' ? clean_(user.idDriver) : clean_(payload.driverId);
  if (!driverId) throw new Error('Pengemudi wajib ditentukan.');

  const driver = findById_(ss.getSheetByName(SHEET.DRIVER), 'Id_Driver', driverId);
  if (!driver) throw new Error('Data driver tidak ditemukan.');

  let vehicle = null;
  let vehicleName = '';
  let kmPerLiter = 0;
  let bbmRutin = 0;
  let jenisBbm = '';

  if (manualVehicle) {
    const noPol = clean_(payload.manualNoPol);
    const merkType = clean_(payload.manualMerkType);
    kmPerLiter = num_(payload.manualKmPerLiter);
    bbmRutin = 0;
    jenisBbm = clean_(payload.manualJenisBbm);
    if (!noPol) throw new Error('No. Pol Kendaraan wajib diisi.');
    if (!merkType) throw new Error('Merk dan Type Kendaraan wajib diisi.');
    if (kmPerLiter <= 0) throw new Error('Konsumsi Km/Liter kendaraan manual harus lebih dari 0.');
    vehicleName = noPol + ' - ' + merkType;
  } else {
    vehicle = findById_(ss.getSheetByName(SHEET.KENDARAAN), 'Id_Kendaraan', kendaraanId);
    if (!vehicle) throw new Error('Data kendaraan tidak ditemukan.');
    kmPerLiter = num_(vehicle.Jumlah_Km_Per_Liter);
    bbmRutin = num_(vehicle.BBM_Rutin);
    jenisBbm = normalizeFuel_(vehicle.Jenis_BBM);
    vehicleName = String(vehicle.No_Pol) + ' - ' + String(vehicle.Merk_Type);
  }

  let destination = null;
  let destinationName = '';
  let jarakPP = 0;

  if (manualDestination) {
    destinationName = clean_(payload.manualDestinationDescription);
    jarakPP = num_(payload.manualDistancePP);
    if (!destinationName) throw new Error('Deskripsi Tujuan wajib diisi.');
    if (jarakPP <= 0) throw new Error('Jarak PP Tujuan harus lebih dari 0.');
  } else {
    destination = findById_(ss.getSheetByName(SHEET.TUJUAN), 'Id_Tujuan', tujuanId);
    if (!destination) throw new Error('Data tujuan tidak ditemukan.');
    destinationName = String(destination.Nama_Tujuan);
    jarakPP = num_(destination.Jarak_PP);
  }

  if (kmPerLiter <= 0) throw new Error('Jumlah Km/Liter kendaraan harus lebih dari 0.');

  let dukungan = (jarakPP / kmPerLiter) - bbmRutin;
  if (APP.CLAMP_NEGATIVE_SUPPORT_TO_ZERO) dukungan = Math.max(0, dukungan);
  dukungan = round2_(dukungan);

  let fileId = '';
  let fileName = '';
  if (payload.suratTugas && payload.suratTugas.dataUrl) {
    const uploaded = saveUploadedFile_(payload.suratTugas, tanggal, driverId, destinationName);
    fileId = uploaded.id;
    fileName = uploaded.name;
  }

  const sh = ss.getSheetByName(SHEET.LAPORAN);
  const id = nextId_(sh, 'Id_Laporan', 'LAP');
  const now = new Date();
  const row = [
    id,
    tanggal,
    String(driver.Id_Driver),
    String(driver.Nama_Lengkap),
    uraian,
    manualVehicle ? '' : String(vehicle.Id_Kendaraan),
    vehicleName,
    manualDestination ? '' : String(destination.Id_Tujuan),
    destinationName,
    jarakPP,
    kmPerLiter,
    bbmRutin,
    jenisBbm,
    dukungan,
    fileId,
    fileName,
    now
  ];
  sh.appendRow(row);

  return {
    ok: true,
    id: id,
    hariTanggal: formatDateValue_(tanggal),
    namaDriver: String(driver.Nama_Lengkap),
    kendaraan: vehicleName,
    tujuan: destinationName,
    jarakPP: jarakPP,
    kmPerLiter: kmPerLiter,
    bbmRutin: bbmRutin,
    jenisBbm: jenisBbm || '-',
    dukungan: dukungan,
    inputRole: user.role,
    inputUsername: user.username
  };
}

// Wrapper with correct implementation (kept separate to keep function name stable).
function getReportList(token, filters) {
  const user = requireAuth_(token, ['DRIVER', 'ADMIN']);
  filters = filters || {};
  let rows = getReportObjects_(getDb_()).map(reportView_);

  if (user.role === 'DRIVER') {
    rows = rows.filter(function(r) { return String(r.idDriver) === String(user.idDriver); });
  }
  if (filters.from) rows = rows.filter(function(r) { return r.hariTanggal >= filters.from; });
  if (filters.to) rows = rows.filter(function(r) { return r.hariTanggal <= filters.to; });
  if (filters.jenisBbm && filters.jenisBbm !== 'SEMUA') {
    rows = rows.filter(function(r) { return normalizeFuel_(r.jenisBbm) === normalizeFuel_(filters.jenisBbm); });
  }
  if (filters.driverId && filters.driverId !== 'SEMUA') {
    rows = rows.filter(function(r) { return String(r.idDriver) === String(filters.driverId); });
  }
  if (filters.kendaraanId && filters.kendaraanId !== 'SEMUA') {
    rows = rows.filter(function(r) { return String(r.idKendaraan) === String(filters.kendaraanId); });
  }
  if (filters.search) {
    const q = String(filters.search).toLowerCase();
    rows = rows.filter(function(r) {
      return [r.hariTanggal, r.uraian, r.kendaraan, r.tujuan, r.namaDriver, r.jenisBbm]
        .join(' ').toLowerCase().indexOf(q) >= 0;
    });
  }

  rows.sort(function(a,b) {
    return String(b.createdAt).localeCompare(String(a.createdAt));
  });
  rows.forEach(function(r, i) { r.no = i + 1; });

  const summary = rows.reduce(function(acc, r) {
    acc.jumlah += 1;
    acc.jarak += num_(r.jarakPP);
    if (normalizeFuel_(r.jenisBbm) === 'PERTALITE') acc.pertalite += num_(r.dukunganBBM);
    if (normalizeFuel_(r.jenisBbm) === 'SOLAR') acc.solar += num_(r.dukunganBBM);
    acc.total += num_(r.dukunganBBM);
    return acc;
  }, { jumlah: 0, jarak: 0, pertalite: 0, solar: 0, total: 0 });

  summary.jarak = round2_(summary.jarak);
  summary.pertalite = round2_(summary.pertalite);
  summary.solar = round2_(summary.solar);
  summary.total = round2_(summary.total);
  return { rows: rows, summary: summary };
}

function deleteReport(token, idLaporan) {
  requireAuth_(token, ['ADMIN']);
  const ss = getDb_();
  const sh = ss.getSheetByName(SHEET.LAPORAN);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(idLaporan)) {
      const fileId = String(data[i][14] || '');
      if (fileId) {
        try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {}
      }
      sh.deleteRow(i + 1);
      return true;
    }
  }
  throw new Error('Laporan tidak ditemukan.');
}

// =============================
// MASTER DATA
// =============================
function getMasterData(token, type) {
  requireAuth_(token, ['ADMIN']);
  const ss = getDb_();
  if (type === 'DRIVER') return readObjects_(ss.getSheetByName(SHEET.DRIVER));
  if (type === 'KENDARAAN') return readObjects_(ss.getSheetByName(SHEET.KENDARAAN));
  if (type === 'TUJUAN') return readObjects_(ss.getSheetByName(SHEET.TUJUAN));
  throw new Error('Tipe master data tidak dikenal.');
}

function getHeaderColumn_(sh, headerName) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const idx = headers.findIndex(function(h) {
    return h.trim().toLowerCase() === String(headerName).trim().toLowerCase();
  });
  if (idx < 0) throw new Error('Kolom ' + headerName + ' tidak ditemukan pada sheet ' + sh.getName() + '.');
  return idx + 1;
}

function normalizeDriverStatus_(value) {
  const s = clean_(value).toUpperCase();
  if (s === 'TIDAK AKTIF' || s === 'NONAKTIF' || s === 'NON-AKTIF') return 'Tidak Aktif';
  return 'Aktif';
}

function saveMasterData(token, type, data) {
  requireAuth_(token, ['ADMIN']);
  data = data || {};
  const ss = getDb_();

  if (type === 'DRIVER') return saveDriver_(ss, data);
  if (type === 'KENDARAAN') return saveVehicle_(ss, data);
  if (type === 'TUJUAN') return saveDestination_(ss, data);
  throw new Error('Tipe master data tidak dikenal.');
}

function saveDriver_(ss, data) {
  const sh = ss.getSheetByName(SHEET.DRIVER);
  ensureDriverStatusColumn_(ss);

  const id = clean_(data.id) || nextId_(sh, 'Id_Driver', 'D');
  const nama = clean_(data.nama);
  const pangkat = clean_(data.pangkat);
  const nrp = clean_(data.nrp);
  const statusInput = clean_(data.status);

  if (!nama || !nrp) throw new Error('Nama lengkap dan NIP/NRP wajib diisi.');

  const existingRow = findRowById_(sh, 1, id);
  const statusCol = getHeaderColumn_(sh, 'Status');
  const existingStatus = existingRow
    ? clean_(sh.getRange(existingRow, statusCol).getValue())
    : '';

  // Driver baru selalu aktif. Saat edit, status lama dipertahankan
  // kecuali status dikirim secara eksplisit.
  const status = statusInput
    ? normalizeDriverStatus_(statusInput)
    : (existingStatus ? normalizeDriverStatus_(existingStatus) : 'Aktif');

  const row = [id, nama, pangkat, nrp, status];

  if (existingRow) {
    sh.getRange(existingRow, 1, 1, row.length).setValues([row]);
  } else {
    sh.appendRow(row);
  }

  return { ok: true, id: id, status: status };
}

function saveVehicle_(ss, data) {
  const sh = ss.getSheetByName(SHEET.KENDARAAN);
  const id = clean_(data.id) || nextId_(sh, 'Id_Kendaraan', 'KND');
  const noPol = clean_(data.noPol);
  const merk = clean_(data.merkType);
  const km = num_(data.kmPerLiter);
  const rutin = num_(data.bbmRutin);
  const fuel = normalizeFuel_(data.jenisBbm);
  if (!noPol || !merk) throw new Error('No. Pol dan Merk/Type wajib diisi.');
  if (km <= 0) throw new Error('Jumlah Km/Liter harus lebih dari 0.');
  if (rutin < 0) throw new Error('BBM Rutin tidak boleh negatif.');
  if (['PERTALITE', 'SOLAR'].indexOf(fuel) < 0) throw new Error('Jenis BBM harus Pertalite atau Solar.');

  const existingRow = findRowById_(sh, 1, id);
  const row = [id, noPol, merk, km, rutin, fuel === 'PERTALITE' ? 'Pertalite' : 'Solar'];
  if (existingRow) sh.getRange(existingRow, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
  return { ok: true, id: id };
}

function saveDestination_(ss, data) {
  const sh = ss.getSheetByName(SHEET.TUJUAN);
  const id = clean_(data.id) || nextId_(sh, 'Id_Tujuan', 'TJN');
  const nama = clean_(data.nama);
  const jarak = num_(data.jarakPP);
  if (!nama) throw new Error('Nama tujuan wajib diisi.');
  if (jarak < 0) throw new Error('Jarak PP tidak boleh negatif.');

  const existingRow = findRowById_(sh, 1, id);
  const row = [id, nama, jarak];
  if (existingRow) sh.getRange(existingRow, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
  return { ok: true, id: id };
}

function deleteMasterData(token, type, id) {
  requireAuth_(token, ['ADMIN']);
  const ss = getDb_();
  let sheetName, idCol, reportCol;
  if (type === 'DRIVER') { sheetName = SHEET.DRIVER; idCol = 'Id_Driver'; reportCol = 2; }
  else if (type === 'KENDARAAN') { sheetName = SHEET.KENDARAAN; idCol = 'Id_Kendaraan'; reportCol = 5; }
  else if (type === 'TUJUAN') { sheetName = SHEET.TUJUAN; idCol = 'Id_Tujuan'; reportCol = 7; }
  else throw new Error('Tipe master data tidak dikenal.');

  const reportRows = getReportObjects_(ss);
  const used = reportRows.some(function(r) {
    if (type === 'DRIVER') return String(r.Id_Driver) === String(id);
    if (type === 'KENDARAAN') return String(r.Id_Kendaraan) === String(id);
    return String(r.Id_Tujuan) === String(id);
  });
  if (used) throw new Error('Data tidak dapat dihapus karena sudah digunakan pada laporan.');

  const sh = ss.getSheetByName(sheetName);
  const row = findRowById_(sh, 1, id);
  if (!row) throw new Error('Data tidak ditemukan.');

  // Untuk DRIVER, "Hapus" adalah soft delete:
  // data tetap tersimpan, hanya status berubah menjadi Tidak Aktif.
  if (type === 'DRIVER') {
    ensureDriverStatusColumn_(ss);
    const statusCol = getHeaderColumn_(sh, 'Status');
    sh.getRange(row, statusCol).setValue('Tidak Aktif');
    return { ok: true, id: id, status: 'Tidak Aktif' };
  }

  // Kendaraan/Tujuan tetap menggunakan perilaku hapus fisik.
  sh.deleteRow(row);
  return true;
}

// =============================
// EXPORT EXCEL
// =============================
function exportExcel(token, filters) {
  requireAuth_(token, ['DRIVER', 'ADMIN']);
  const result = getReportList(token, filters || {});
  const ssTemp = SpreadsheetApp.create('EXPORT_BBM_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss'));
  const sh = ssTemp.getSheets()[0];
  sh.setName('Laporan BBM');

  const headers = ['No', 'Hari & Tanggal', 'Uraian Kegiatan', 'Kendaraan', 'Tujuan', 'Jarak Tempuh (PP)', 'Dukungan BBM', 'Jenis BBM', 'Surat Tugas'];
  const values = [headers].concat(result.rows.map(function(r) {
    return [
      r.no,
      r.hariTanggal,
      r.uraian,
      r.kendaraan,
      r.tujuan,
      r.jarakPP,
      r.dukunganBBM,
      r.jenisBbm,
      r.suratTugasNama || ''
    ];
  }));

  sh.getRange(1, 1, values.length, headers.length).setValues(values);
  formatHeader_(sh, headers.length);
  sh.autoResizeColumns(1, headers.length);
  sh.getDataRange().setVerticalAlignment('middle');

  const exportBlob = UrlFetchApp.fetch(
    'https://docs.google.com/spreadsheets/d/' + ssTemp.getId() + '/export?format=xlsx',
    { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true }
  ).getBlob();

  try { DriveApp.getFileById(ssTemp.getId()).setTrashed(true); } catch (e) {}

  if (!exportBlob || exportBlob.getBytes().length < 100) {
    throw new Error('Export Excel gagal dibuat. Silakan coba kembali.');
  }
  const filename = 'Laporan_BBM_'
    + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '.xlsx';
  const bytes = exportBlob.getBytes();
  return {
    filename: filename,
    mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    base64: Utilities.base64Encode(bytes),
    totalRows: result.rows.length
  };
}

// =============================
// FILE SURAT TUGAS
// =============================
function saveUploadedFile_(file, tanggal, driverId, tujuan) {
  const dataUrl = String(file.dataUrl || '');
  const match = dataUrl.match(/^data:([^;]+);base64,(.+)$/);
  if (!match) throw new Error('Format file Surat Tugas tidak valid.');
  const mime = match[1];
  const base64 = match[2];
  const bytes = Utilities.base64Decode(base64);
  if (bytes.length > APP.MAX_UPLOAD_BYTES) throw new Error('Ukuran Surat Tugas maksimal 5 MB.');

  const props = PropertiesService.getScriptProperties();
  const folderId = props.getProperty(APP.SURAT_FOLDER_PROP) || APP.FIXED_SURAT_FOLDER_ID;
  props.setProperty(APP.SURAT_FOLDER_PROP, folderId);
  let folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (e) {
    throw new Error('Folder Surat Tugas tidak dapat diakses. Pastikan project/deployment memiliki akses ke folder Drive tersebut. Detail: ' + e.message);
  }

  const safeOriginal = String(file.name || 'surat_tugas').replace(/[^a-zA-Z0-9._-]+/g, '_');
  const datePart = String(tanggal).replace(/-/g, '');
  const name = datePart + '_' + driverId + '_' + String(tujuan).replace(/[^a-zA-Z0-9_-]+/g, '_') + '_' + safeOriginal;
  const blob = Utilities.newBlob(bytes, mime, name);
  const created = folder.createFile(blob);
  return { id: created.getId(), name: created.getName() };
}

function getFileData(token, fileId) {
  requireAuth_(token, ['DRIVER', 'ADMIN']);
  if (!fileId) throw new Error('File tidak ditemukan.');
  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();
  if (blob.getBytes().length > APP.MAX_UPLOAD_BYTES) throw new Error('File terlalu besar untuk pratinjau.');
  return {
    name: file.getName(),
    mime: blob.getContentType(),
    base64: Utilities.base64Encode(blob.getBytes())
  };
}

// =============================
// HELPERS
// =============================
function getReportObjects_(ss) {
  if (!ss) ss = getDb_();
  return readObjects_(ss.getSheetByName(SHEET.LAPORAN));
}

function reportView_(r) {
  return {
    id: String(r.Id_Laporan || ''),
    no: 0,
    hariTanggal: formatDateValue_(r.Hari_Tanggal),
    namaDriver: String(r.Nama_Driver || ''),
    idDriver: String(r.Id_Driver || ''),
    uraian: String(r.Uraian_Kegiatan || ''),
    kendaraan: String(r.Kendaraan || ''),
    idKendaraan: String(r.Id_Kendaraan || ''),
    tujuan: String(r.Tujuan || ''),
    idTujuan: String(r.Id_Tujuan || ''),
    jarakPP: num_(r.Jarak_PP),
    kmPerLiter: num_(r.Km_Per_Liter),
    bbmRutin: num_(r.BBM_Rutin),
    jenisBbm: normalizeFuel_(r.Jenis_BBM),
    dukunganBBM: round2_(num_(r.Dukungan_BBM)),
    suratTugasFileId: String(r.Surat_Tugas_File_Id || ''),
    suratTugasNama: String(r.Surat_Tugas_Nama || ''),
    createdAt: String(r.Created_At || '')
  };
}

function readObjects_(sh) {
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(String);
  return values.slice(1).filter(function(row) {
    return row.some(function(cell) { return String(cell).trim() !== ''; });
  }).map(function(row) {
    const obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    return obj;
  });
}

function findById_(sh, fieldName, id) {
  const rows = readObjects_(sh);
  return rows.find(function(r) { return String(r[fieldName]) === String(id); }) || null;
}

function findRowById_(sh, idColIndex, id) {
  if (sh.getLastRow() < 2) return null;
  const ids = sh.getRange(2, idColIndex, sh.getLastRow() - 1, 1).getValues().map(function(r) { return String(r[0]); });
  const idx = ids.findIndex(function(v) { return v === String(id); });
  return idx < 0 ? null : idx + 2;
}

function nextId_(sh, fieldName, prefix) {
  const rows = readObjects_(sh);
  let max = 0;
  rows.forEach(function(r) {
    const value = String(r[fieldName] || '');
    const match = value.match(/(\d+)$/);
    if (match) max = Math.max(max, Number(match[1]));
  });
  return prefix + String(max + 1).padStart(3, '0');
}

function validateDateInput_(value) {
  const s = clean_(value);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) throw new Error('Tanggal tidak valid.');
  return s;
}

function formatDateValue_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const s = String(value || '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  // fallback for sheet date strings
  const d = new Date(s);
  return isNaN(d) ? s : Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function normalizeFuel_(value) {
  const s = clean_(value).toUpperCase();
  if (s.indexOf('SOLAR') >= 0) return 'SOLAR';
  return 'PERTALITE';
}

function clean_(value) {
  return value == null ? '' : String(value).trim();
}

function num_(value) {
  if (typeof value === 'number') return isFinite(value) ? value : 0;
  const s = String(value == null ? '' : value).replace(/\s/g, '').replace(',', '.');
  const n = Number(s);
  return isFinite(n) ? n : 0;
}

function round2_(value) {
  return Math.round(num_(value) * 100) / 100;
}
