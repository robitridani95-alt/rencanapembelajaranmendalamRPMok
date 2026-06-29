// ============================================================
// PEMBELAJARAN MENDALAM - DOWNLOAD LINK SYSTEM
// Google Apps Script Backend (Pure API — no HTML serve)
// v3.0 — Database only, dipanggil via fetch dari HTML external
// ============================================================

const SHEETS = {
  ADMIN: 'Admin',
  GURU: 'Guru',
  MATERI: 'Materi',
  LOG: 'LogDownload'
};

const JENIS_DOKUMEN = ['Modul Ajar', 'CP', 'ATP', 'PROSEM', 'PROTA', 'KKTP', 'Power Point', 'LKPD'];

// ============================================================
// ENTRY POINT — Handle GET & POST dengan CORS
// ============================================================
function doGet(e) {
  const action = e && e.parameter && e.parameter.action ? e.parameter.action : '';
  const data   = e && e.parameter && e.parameter.data   ? JSON.parse(e.parameter.data) : {};
  const result = action ? processRequest(action, data) : { status: 'ok', message: 'API Pembelajaran Mendalam v3.0 aktif' };
  return buildResponse(result);
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch(err) {}
  const result = processRequest(body.action || '', body.data || {});
  return buildResponse(result);
}

function buildResponse(result) {
  const output = ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ============================================================
// ROUTER
// ============================================================
function processRequest(action, data) {
  try {
    switch(action) {
      case 'init':               return initSheets();
      case 'login':              return login(data);
      case 'loginGuru':          return loginGuru(data);
      case 'daftarGuru':         return daftarGuru(data);
      case 'getAllGuru':          return getAllGuru();
      case 'aktivasiGuru':       return aktivasiGuru(data);
      case 'tolakGuru':          return tolakGuru(data);
      case 'hapusGuru':          return hapusGuru(data);
      case 'getMateri':          return getMateri(data);
      case 'addMateri':          return addMateri(data);
      case 'editMateri':         return editMateri(data);
      case 'hapusMateri':        return hapusMateri(data);
      case 'downloadMateri':     return downloadMateri(data);
      case 'getLog':             return getLog(data);
      case 'getDashboard':       return getDashboard();
      case 'getMapelByJenjang':  return getMapelByJenjang(data);
      case 'getJenisDokumen':    return { status: 'ok', data: JENIS_DOKUMEN };
      case 'changePasswordGuru': return changePasswordGuru(data);
      default: return { status: 'error', message: 'Aksi tidak dikenal: ' + action };
    }
  } catch(err) {
    return { status: 'error', message: err.toString() };
  }
}

// ============================================================
// INISIALISASI SHEET
// ============================================================
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let adminSheet = ss.getSheetByName(SHEETS.ADMIN);
  if (!adminSheet) {
    adminSheet = ss.insertSheet(SHEETS.ADMIN);
    adminSheet.appendRow(['ID','Nama','Username','Password','Email','TanggalDibuat']);
    adminSheet.appendRow(['ADM001','Super Admin','admin',hashPassword('admin123'),'admin@sekolah.id',new Date().toISOString()]);
    formatHeader(adminSheet);
  }

  let guruSheet = ss.getSheetByName(SHEETS.GURU);
  if (!guruSheet) {
    guruSheet = ss.insertSheet(SHEETS.GURU);
    guruSheet.appendRow(['ID','Nama','NIP','Email','MataPelajaran','Jenjang','Sekolah','Username','Password','Status','TanggalDaftar','TanggalAktivasi','AktivasiOleh']);
    formatHeader(guruSheet);
  }

  let materiSheet = ss.getSheetByName(SHEETS.MATERI);
  if (!materiSheet) {
    materiSheet = ss.insertSheet(SHEETS.MATERI);
    materiSheet.appendRow(['ID','Judul','Deskripsi','Jenjang','MataPelajaran','JenisDokumen','Topik','LinkDownload','TipeFile','Ukuran','TanggalUpload','UploadOleh','JumlahDownload','Status']);
    formatHeader(materiSheet);
    seedMateriData(materiSheet);
  } else {
    const headers = materiSheet.getRange(1,1,1,materiSheet.getLastColumn()).getValues()[0];
    if (!headers.includes('JenisDokumen')) {
      materiSheet.insertColumnAfter(5);
      materiSheet.getRange(1,6).setValue('JenisDokumen');
      const lastRow = materiSheet.getLastRow();
      if (lastRow > 1) materiSheet.getRange(2,6,lastRow-1,1).setValue('Modul Ajar');
    }
  }

  let logSheet = ss.getSheetByName(SHEETS.LOG);
  if (!logSheet) {
    logSheet = ss.insertSheet(SHEETS.LOG);
    logSheet.appendRow(['ID','GuruID','NamaGuru','MateriID','JudulMateri','Jenjang','MataPelajaran','JenisDokumen','TanggalDownload','IP']);
    formatHeader(logSheet);
  }

  return { status: 'ok', message: 'Spreadsheet siap' };
}

function formatHeader(sheet) {
  const r = sheet.getRange(1,1,1,sheet.getLastColumn());
  r.setBackground('#1e3a5f'); r.setFontColor('#ffffff'); r.setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function seedMateriData(sheet) {
  const MAPEL = {
    SD:  ['PPKn','Bahasa Indonesia','Matematika','IPAS','Seni Budaya','PJOK','Bahasa Inggris','Pendidikan Agama Islam'],
    SMP: ['PPKn','Bahasa Indonesia','Matematika','IPA','IPS','Seni Budaya','PJOK','Bahasa Inggris','Informatika','Pendidikan Agama Islam'],
    SMA: ['PPKn','Bahasa Indonesia','Matematika','Fisika','Kimia','Biologi','Sejarah Indonesia','Bahasa Inggris','Informatika','Ekonomi','Sosiologi','Pendidikan Agama Islam'],
    SMK: ['PPKn','Bahasa Indonesia','Matematika','Bahasa Inggris','PJOK','Informatika','Dasar-Dasar Kejuruan','Pendidikan Agama Islam']
  };
  const TOPIK = ['Pengalaman Belajar Konkret','Refleksi dan Metakognisi','Berpikir Kritis dan Analitis'];
  let id = 1;
  Object.entries(MAPEL).forEach(([jenjang, mapelList]) => {
    mapelList.forEach(mp => {
      JENIS_DOKUMEN.forEach(jenis => {
        const topik = jenis === 'Modul Ajar' ? TOPIK[id % 3] : '-';
        sheet.appendRow([
          'MAT' + String(id).padStart(4,'0'),
          `${jenis}: ${mp} ${jenjang}`,
          `${jenis} mata pelajaran ${mp} jenjang ${jenjang} sesuai Kurikulum Merdeka & Permendikdasmen No. 13/2025.`,
          jenjang, mp, jenis, topik,
          'https://drive.google.com/file/d/EXAMPLE_' + id + '/view',
          jenis==='Power Point'?'PPTX':'PDF', Math.floor(Math.random()*4+1)+' MB',
          new Date().toISOString(), 'Admin', 0, 'Aktif'
        ]);
        id++;
      });
    });
  });
}

// ============================================================
// HELPER
// ============================================================
function hashPassword(pass) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pass)
    .map(b => ('0'+(b&0xFF).toString(16)).slice(-2)).join('');
}
function generateId(prefix) { return prefix + new Date().getTime(); }
function getSheet(name)      { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }
function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h,i) => obj[h] = row[i]);
    return obj;
  });
}

// ============================================================
// AUTH
// ============================================================
function login(data) {
  const rows = sheetToObjects(getSheet(SHEETS.ADMIN));
  const hashed = hashPassword(data.password);
  const admin = rows.find(r => r.Username===data.username && r.Password===hashed);
  if (admin) return { status:'ok', role:'admin', nama:admin.Nama, id:admin.ID };
  return { status:'error', message:'Username atau password salah' };
}

function loginGuru(data) {
  const rows = sheetToObjects(getSheet(SHEETS.GURU));
  const hashed = hashPassword(data.password);
  const guru = rows.find(r => r.Username===data.username && r.Password===hashed);
  if (!guru) return { status:'error', message:'Username atau password salah' };
  if (guru.Status==='Menunggu') return { status:'error', message:'Akun menunggu aktivasi Admin' };
  if (guru.Status==='Ditolak')  return { status:'error', message:'Akun Anda ditolak Admin' };
  if (guru.Status!=='Aktif')    return { status:'error', message:'Akun tidak aktif' };
  return { status:'ok', role:'guru', nama:guru.Nama, id:guru.ID, nip:guru.NIP, mapel:guru.MataPelajaran, jenjang:guru.Jenjang, sekolah:guru.Sekolah };
}

// ============================================================
// GURU
// ============================================================
function daftarGuru(data) {
  const sheet = getSheet(SHEETS.GURU);
  const rows  = sheetToObjects(sheet);
  if (rows.find(r => r.Username===data.username)) return { status:'error', message:'Username sudah digunakan' };
  if (data.nip && rows.find(r => r.NIP===data.nip)) return { status:'error', message:'NIP sudah terdaftar' };
  sheet.appendRow([generateId('GRU'),data.nama,data.nip,data.email,data.mapel,data.jenjang,data.sekolah,data.username,hashPassword(data.password),'Menunggu',new Date().toISOString(),'','']);
  return { status:'ok', message:'Pendaftaran berhasil! Tunggu aktivasi Admin.' };
}
function getAllGuru() {
  const rows = sheetToObjects(getSheet(SHEETS.GURU));
  return { status:'ok', data: rows.map(r => ({...r, Password:undefined})) };
}
function aktivasiGuru(data) {
  return updateGuruField(data.guruId, [
    {field:'Status',val:'Aktif'},{field:'TanggalAktivasi',val:new Date().toISOString()},{field:'AktivasiOleh',val:data.adminNama}
  ], 'Guru diaktivasi');
}
function tolakGuru(data)  { return updateGuruField(data.guruId, [{field:'Status',val:'Ditolak'}], 'Guru ditolak'); }
function hapusGuru(data) {
  const sheet = getSheet(SHEETS.GURU);
  const allData = sheet.getDataRange().getValues();
  const idIdx = allData[0].indexOf('ID');
  for (let i=1;i<allData.length;i++) {
    if (allData[i][idIdx]===data.guruId) { sheet.deleteRow(i+1); return {status:'ok',message:'Data guru dihapus'}; }
  }
  return {status:'error',message:'Guru tidak ditemukan'};
}
function updateGuruField(guruId, updates, successMsg) {
  const sheet = getSheet(SHEETS.GURU);
  const allData = sheet.getDataRange().getValues();
  const headers = allData[0];
  const idIdx = headers.indexOf('ID');
  for (let i=1;i<allData.length;i++) {
    if (allData[i][idIdx]===guruId) {
      updates.forEach(u => { const idx=headers.indexOf(u.field); if(idx>=0) sheet.getRange(i+1,idx+1).setValue(u.val); });
      return {status:'ok',message:successMsg};
    }
  }
  return {status:'error',message:'Guru tidak ditemukan'};
}
function changePasswordGuru(data) {
  const sheet = getSheet(SHEETS.GURU);
  const allData = sheet.getDataRange().getValues();
  const headers = allData[0];
  const idIdx   = headers.indexOf('ID');
  const passIdx = headers.indexOf('Password');
  for (let i=1;i<allData.length;i++) {
    if (allData[i][idIdx]===data.guruId) {
      if (allData[i][passIdx]!==hashPassword(data.oldPassword)) return {status:'error',message:'Password lama tidak sesuai'};
      sheet.getRange(i+1,passIdx+1).setValue(hashPassword(data.newPassword));
      return {status:'ok',message:'Password berhasil diubah'};
    }
  }
  return {status:'error',message:'Guru tidak ditemukan'};
}

// ============================================================
// MATA PELAJARAN
// ============================================================
const MAPEL_BY_JENJANG = {
  SD:  ['Pendidikan Agama Islam','Pendidikan Agama Kristen','Pendidikan Agama Katolik','Pendidikan Agama Hindu','Pendidikan Agama Buddha','Pendidikan Agama Konghucu','PPKn','Bahasa Indonesia','Matematika','IPAS','Seni Budaya','PJOK','Bahasa Inggris','Muatan Lokal'],
  SMP: ['Pendidikan Agama Islam','Pendidikan Agama Kristen','Pendidikan Agama Katolik','Pendidikan Agama Hindu','Pendidikan Agama Buddha','Pendidikan Agama Konghucu','PPKn','Bahasa Indonesia','Matematika','IPA','IPS','Seni Budaya','PJOK','Bahasa Inggris','Informatika','Muatan Lokal'],
  SMA: ['Pendidikan Agama Islam','Pendidikan Agama Kristen','Pendidikan Agama Katolik','Pendidikan Agama Hindu','Pendidikan Agama Buddha','Pendidikan Agama Konghucu','PPKn','Bahasa Indonesia','Matematika','Fisika','Kimia','Biologi','Sejarah Indonesia','Seni Budaya','PJOK','Bahasa Inggris','Informatika','Ekonomi','Sosiologi','Geografi','Bahasa Jepang','Bahasa Jerman','Bahasa Arab','Bahasa Mandarin','Antropologi','Muatan Lokal'],
  SMK: ['Pendidikan Agama Islam','Pendidikan Agama Kristen','Pendidikan Agama Katolik','Pendidikan Agama Hindu','Pendidikan Agama Buddha','Pendidikan Agama Konghucu','PPKn','Bahasa Indonesia','Matematika','Bahasa Inggris','PJOK','Informatika','Projek Ilmu Pengetahuan Alam dan Sosial','Dasar-Dasar Kejuruan','Konsentrasi Keahlian','Muatan Lokal']
};
function getMapelByJenjang(data) { return {status:'ok', data: MAPEL_BY_JENJANG[data.jenjang]||[]}; }

// ============================================================
// MATERI
// ============================================================
function getMateri(data) {
  const rows = sheetToObjects(getSheet(SHEETS.MATERI));
  let f = rows.filter(r => r.Status==='Aktif');
  if (data && data.jenjang) f = f.filter(r => r.Jenjang===data.jenjang);
  if (data && data.mapel)   f = f.filter(r => r.MataPelajaran===data.mapel);
  if (data && data.jenis)   f = f.filter(r => r.JenisDokumen===data.jenis);
  if (data && data.topik)   f = f.filter(r => r.Topik===data.topik);
  if (data && data.search) {
    const q = data.search.toLowerCase();
    f = f.filter(r => (r.Judul||'').toLowerCase().includes(q)||(r.Deskripsi||'').toLowerCase().includes(q)||(r.MataPelajaran||'').toLowerCase().includes(q));
  }
  return {status:'ok', data:f};
}
function addMateri(data) {
  const sheet = getSheet(SHEETS.MATERI);
  const id = generateId('MAT');
  sheet.appendRow([id,data.judul,data.deskripsi,data.jenjang,data.mapel,data.jenis||'Modul Ajar',data.topik||'-',data.link,data.tipe||'PDF',data.ukuran||'-',new Date().toISOString(),data.uploadOleh,0,'Aktif']);
  return {status:'ok', message:'Materi berhasil ditambahkan', id};
}
function editMateri(data) {
  const sheet = getSheet(SHEETS.MATERI);
  const allData = sheet.getDataRange().getValues();
  const headers = allData[0];
  const idIdx = headers.indexOf('ID');
  for (let i=1;i<allData.length;i++) {
    if (allData[i][idIdx]===data.id) {
      const cols = {Judul:data.judul,Deskripsi:data.deskripsi,Jenjang:data.jenjang,MataPelajaran:data.mapel,JenisDokumen:data.jenis,Topik:data.topik,LinkDownload:data.link,TipeFile:data.tipe,Ukuran:data.ukuran,Status:data.status};
      Object.entries(cols).forEach(([key,val]) => { const idx=headers.indexOf(key); if(idx>=0&&val!==undefined) sheet.getRange(i+1,idx+1).setValue(val); });
      return {status:'ok',message:'Materi diupdate'};
    }
  }
  return {status:'error',message:'Materi tidak ditemukan'};
}
function hapusMateri(data) {
  const sheet = getSheet(SHEETS.MATERI);
  const allData = sheet.getDataRange().getValues();
  const headers = allData[0];
  const idIdx = headers.indexOf('ID');
  const stIdx = headers.indexOf('Status');
  for (let i=1;i<allData.length;i++) {
    if (allData[i][idIdx]===data.id) { sheet.getRange(i+1,stIdx+1).setValue('Nonaktif'); return {status:'ok',message:'Materi dihapus'}; }
  }
  return {status:'error',message:'Materi tidak ditemukan'};
}
function downloadMateri(data) {
  const logSheet = getSheet(SHEETS.LOG);
  logSheet.appendRow([generateId('LOG'),data.guruId,data.namaGuru,data.materiId,data.judulMateri,data.jenjang,data.mapel,data.jenis||'-',new Date().toISOString(),'']);
  const sheet = getSheet(SHEETS.MATERI);
  const allData = sheet.getDataRange().getValues();
  const headers = allData[0];
  const idIdx  = headers.indexOf('ID');
  const dlIdx  = headers.indexOf('JumlahDownload');
  const linkIdx = headers.indexOf('LinkDownload');
  for (let i=1;i<allData.length;i++) {
    if (allData[i][idIdx]===data.materiId) {
      sheet.getRange(i+1,dlIdx+1).setValue((parseInt(allData[i][dlIdx])||0)+1);
      return {status:'ok', link:allData[i][linkIdx]};
    }
  }
  return {status:'error',message:'Materi tidak ditemukan'};
}

// ============================================================
// LOG
// ============================================================
function getLog(data) {
  const rows = sheetToObjects(getSheet(SHEETS.LOG));
  let f = rows;
  if (data && data.guruId) f = f.filter(r => r.GuruID===data.guruId);
  return {status:'ok', data: f.reverse().slice(0,100)};
}

// ============================================================
// DASHBOARD
// ============================================================
function getDashboard() {
  const guru   = sheetToObjects(getSheet(SHEETS.GURU));
  const materi = sheetToObjects(getSheet(SHEETS.MATERI));
  const log    = sheetToObjects(getSheet(SHEETS.LOG));
  const aktif  = materi.filter(m => m.Status==='Aktif');
  const byJenjang = {};
  ['SD','SMP','SMA','SMK'].forEach(j => byJenjang[j] = aktif.filter(m=>m.Jenjang===j).length);
  const byJenis = {};
  JENIS_DOKUMEN.forEach(jn => byJenis[jn] = aktif.filter(m=>m.JenisDokumen===jn).length);
  return {
    status:'ok',
    data: {
      totalGuru:     guru.filter(g=>g.Status==='Aktif').length,
      guruMenunggu:  guru.filter(g=>g.Status==='Menunggu').length,
      totalMateri:   aktif.length,
      totalDownload: log.length,
      byJenjang, byJenis,
      recentLog: log.reverse().slice(0,5)
    }
  };
}
