/**
 * Fahdking Datsis - DASHBOARD GURU
 * Versi Optimized - Anti Lag HP
 * Perubahan utama:
 * - getAllInitialData() HANYA memuat data ringan (kelas & setting)
 * - Tambah fungsi getSiswaByKelas() untuk lazy load siswa per kelas
 * - Tambah fungsi getJadwalData() untuk lazy load jadwal
 * - Fungsi lain tetap sama
 */

// ==========================================
// 1. FUNGSI UTAMA UNTUK MENAMPILKAN WEB
// ==========================================

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Fahdking Datsis - Dashboard Guru')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- FITUR LOGIN ---
function cekLogin(username, password) {
  const USER_VALID = "admin";
  const PASS_VALID = "Admin123";
  if (username === USER_VALID && password === PASS_VALID) {
    return { success: true, message: "Login berhasil!" };
  } else {
    return { success: false, message: "Username atau Password salah!" };
  }
}

// ==========================================
// 2. FUNGSI SETUP DATABASE
// ==========================================

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = {
    'Kelas': ['Nama Kelas'],
    'Siswa': ['No', 'Nama Siswa', 'Kelas'],
    'Jadwal': ['Hari', 'Jam Mulai', 'Jam Selesai', 'Mata Pelajaran', 'Kelas'],
    'Absensi': ['Timestamp', 'Tanggal', 'Kelas', 'No', 'Nama Siswa', 'Status Kehadiran'],
    'Nilai': ['Timestamp', 'Tanggal', 'Kelas', 'Kategori Nilai', 'No', 'Nama Siswa', 'Nilai Angka', 'Bab', 'Tujuan Pembelajaran', 'Bentuk'],
    'Jurnal': ['Timestamp', 'Tanggal', 'Jam Ke', 'Kelas', 'Mata Pelajaran', 'Materi Pokok', 'Kegiatan Pembelajaran', 'Keterangan'],
    'Setting': ['Nama Sekolah', 'Alamat Lengkap', 'Nama Guru', 'NIP Guru', 'Nama Kepala Sekolah', 'NIP Kepsek']
  };
  for (let sheetName in config) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName);
    let headers = config[sheetName];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
}

// ==========================================
// 3. FUNGSI MEMUAT DATA AWAL (RINGAN - anti lag)
//    Hanya kelas & setting, TIDAK semua siswa
// ==========================================

function getAllInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const getSheetData = (name) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) return [];
    let lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    return sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getDisplayValues();
  };

  // OPTIMASI: hanya muat kelas & setting saat login
  // siswa & jadwal dimuat terpisah saat dibutuhkan
  return {
    kelas: getSheetData('Kelas'),
    setting: getSheetData('Setting')
  };
}

// ==========================================
// 3b. LAZY LOAD: Siswa & Jadwal (dipanggil terpisah)
// ==========================================

function getSiswaData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Siswa');
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues();
}

function getJadwalData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Jadwal');
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues();
}

// Ambil siswa berdasarkan kelas tertentu (lebih efisien untuk absensi/nilai)
function getSiswaByKelas(namaKelas) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Siswa');
  if (!sheet || sheet.getLastRow() < 2) return [];
  const allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues();
  return allData.filter(row => row[2] === namaKelas);
}

// ==========================================
// 4. FUNGSI SIMPAN DATA SATUAN
// ==========================================

function simpanKelas(data) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Kelas').appendRow(data);
  return "Data Kelas ditambahkan!";
}

function simpanSiswa(data) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Siswa').appendRow(data);
  return "Data Siswa ditambahkan!";
}

function simpanJadwal(data) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Jadwal').appendRow(data);
  return "Jadwal disimpan!";
}

function simpanJurnal(data) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Jurnal').appendRow(data);
  return "Jurnal disimpan!";
}

function simpanSetting(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Setting');
  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  sheet.getRange(2, 1, 1, data.length).setValues([data]);
  return "Pengaturan Identitas berhasil disimpan!";
}

// --- FITUR HAPUS DATA ---
function hapusDataBaris(sheetName, rowDataJSON) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return "Sheet tidak ditemukan";
  const targetData = JSON.parse(rowDataJSON);
  const data = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    let sheetRow = data[i].slice(0, targetData.length);
    if (JSON.stringify(sheetRow) === JSON.stringify(targetData)) {
      sheet.deleteRow(i + 1);
      return "Data berhasil dihapus!";
    }
  }
  throw new Error("Data tidak ditemukan di database.");
}

// ==========================================
// 5. FUNGSI SIMPAN DATA MASAL (ABSENSI & NILAI)
// ==========================================

function simpanAbsensiMasal(dataArray) {
  if (dataArray.length === 0) throw new Error("Tidak ada data absensi untuk disimpan.");
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Absensi');
  sheet.getRange(sheet.getLastRow() + 1, 1, dataArray.length, dataArray[0].length).setValues(dataArray);
  return "Data Absensi berhasil disimpan!";
}

function simpanNilaiMasal(dataArray) {
  if (dataArray.length === 0) throw new Error("Tidak ada data nilai untuk disimpan.");
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Nilai');
  sheet.getRange(sheet.getLastRow() + 1, 1, dataArray.length, dataArray[0].length).setValues(dataArray);
  return "Data Nilai berhasil disimpan!";
}

// ==========================================
// 6. LOGIC FILTER & AGREGASI REKAP
// ==========================================

function getDataRekapanHelper(kelas, jenis, tglMulai, tglAkhir) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetData = ss.getSheetByName(jenis);

  let result = { headers: [], data: [], isGrouped: false };
  if (!sheetData || sheetData.getLastRow() < 2) return result;

  let allData = sheetData.getRange(2, 1, sheetData.getLastRow() - 1, sheetData.getLastColumn()).getDisplayValues();

  let filteredData = allData.filter(row => {
    let tgl = row[1];
    let kls = (jenis === 'Jurnal') ? row[3] : row[2];
    let matchKelas = (kelas === "Semua") ? true : (kls === kelas);
    let matchTgl = (tgl >= tglMulai && tgl <= tglAkhir);
    return matchKelas && matchTgl;
  });

  if (jenis === 'Absensi') {
    result.headers = ['No', 'Nama Lengkap', 'Hadir', 'Sakit', 'Izin', 'Alfa'];
    let rekapSiswa = {};
    filteredData.forEach(row => {
      let noSiswa = row[3];
      let namaSiswa = row[4];
      let status = row[5];
      let key = noSiswa + "_" + namaSiswa;
      if (!rekapSiswa[key]) {
        rekapSiswa[key] = { no: noSiswa, nama: namaSiswa, H: 0, S: 0, I: 0, A: 0 };
      }
      if (status === 'Hadir') rekapSiswa[key].H++;
      else if (status === 'Sakit') rekapSiswa[key].S++;
      else if (status === 'Izin') rekapSiswa[key].I++;
      else if (status === 'Alfa') rekapSiswa[key].A++;
    });
    result.data = Object.values(rekapSiswa).map(s => [s.no, s.nama, s.H, s.S, s.I, s.A]);
    result.data.sort((a, b) => parseInt(a[0]) - parseInt(b[0]));

  } else if (jenis === 'Nilai') {
    result.isGrouped = true;
    let groupedData = {};
    filteredData.forEach(row => {
      let tgl = row[1];
      let kategori = row[3];
      let noSiswa = row[4];
      let namaSiswa = row[5];
      let nilai = row[6];
      let bab = row[7] || "-";
      let tujuan = row[8] || "-";
      let bentuk = row[9] || "-";
      let groupKey = tgl + "_" + kategori + "_" + bab + "_" + bentuk;
      if (!groupedData[groupKey]) {
        groupedData[groupKey] = { kategori, tanggal: tgl, bab, tujuan, bentuk, siswa: [] };
      }
      groupedData[groupKey].siswa.push([noSiswa, namaSiswa, nilai]);
    });
    for (let key in groupedData) {
      groupedData[key].siswa.sort((a, b) => parseInt(a[0]) - parseInt(b[0]));
    }
    result.data = Object.values(groupedData);

  } else if (jenis === 'Jurnal') {
    result.headers = ['Tanggal', 'Jam', 'Materi Pokok', 'Kegiatan Pembelajaran', 'Keterangan'];
    result.data = filteredData.map(row => [row[1], row[2], row[5], row[6], row[7]]);
  }

  return result;
}

function getPreviewRekap(kelas, jenis, tglMulai, tglAkhir) {
  return getDataRekapanHelper(kelas, jenis, tglMulai, tglAkhir);
}

function prosesRekapCetak(kelas, jenis, tglMulai, tglAkhir) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetSetting = ss.getSheetByName('Setting');

  let kop = ["NAMA SEKOLAH", "ALAMAT SEKOLAH", "NAMA GURU", "NIP GURU", "NAMA KEPSEK", "NIP KEPSEK"];
  if (sheetSetting && sheetSetting.getLastRow() > 1) {
    kop = sheetSetting.getRange(2, 1, 1, 6).getDisplayValues()[0];
  }

  let rekapData = getDataRekapanHelper(kelas, jenis, tglMulai, tglAkhir);
  let textKelas = (kelas === "Semua") ? "Semua Kelas" : `Kelas ${kelas}`;
  let tabelHtml = "";

  if (rekapData.isGrouped && jenis === 'Nilai') {
    if (rekapData.data.length === 0) {
      tabelHtml = `<p style="text-align:center; font-weight:bold;">Tidak ada data pada rentang tanggal tersebut.</p>`;
    } else {
      // OPTIMASI: gunakan array join, bukan += di loop
      let parts = [];
      rekapData.data.forEach(grup => {
        parts.push(`
          <div class="meta-nilai">
            <table>
              <tr><td width="150px"><b>Jenis</b></td><td>: ${grup.kategori}</td></tr>
              <tr><td><b>Tanggal</b></td><td>: ${grup.tanggal}</td></tr>
              <tr><td><b>Bab</b></td><td>: ${grup.bab}</td></tr>
              <tr><td><b>Tujuan Pembelajaran</b></td><td>: ${grup.tujuan}</td></tr>
              <tr><td><b>Bentuk</b></td><td>: ${grup.bentuk}</td></tr>
            </table>
          </div>
          <table class="data-table">
            <thead><tr><th width="50px">No</th><th>Nama Siswa</th><th width="100px">Nilai</th></tr></thead>
            <tbody>
              ${grup.siswa.map(s => `<tr><td style="text-align:center">${s[0]}</td><td>${s[1]}</td><td style="text-align:center">${s[2]}</td></tr>`).join('')}
            </tbody>
          </table><br>
        `);
      });
      tabelHtml = parts.join('');
    }
  } else {
    let headers = rekapData.headers;
    let filteredData = rekapData.data;
    tabelHtml = `
      <table class="data-table">
        <thead>
          <tr>${headers.map(h => `<th>${h}</th>`).join('')}</tr>
        </thead>
        <tbody>
          ${filteredData.length > 0
            ? filteredData.map(row => `<tr>${row.map(cell => `<td>${cell}</td>`).join('')}</tr>`).join('')
            : `<tr><td colspan="${headers.length}" style="text-align:center;">Tidak ada data pada rentang tanggal tersebut.</td></tr>`
          }
        </tbody>
      </table>
    `;
  }

  let htmlTemplate = `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        body { font-family: 'Helvetica', 'Arial', sans-serif; font-size: 12px; }
        .header { text-align: center; border-bottom: 3px solid black; padding-bottom: 10px; margin-bottom: 20px; }
        .header h1 { margin: 0; font-size: 18px; text-transform: uppercase; }
        .header p { margin: 2px 0; font-size: 12px; }
        h3 { text-align: center; text-transform: uppercase; margin-bottom: 5px; }
        .periode { text-align: center; margin-bottom: 20px; font-style: italic; }
        .meta-nilai table { width: 100%; border: none; margin-bottom: 5px; }
        .meta-nilai td { border: none; padding: 2px; text-align: left; }
        .data-table { width: 100%; border-collapse: collapse; margin-bottom: 30px; }
        .data-table th, .data-table td { border: 1px solid #000; padding: 6px 8px; text-align: left; }
        .data-table th { background-color: #f2f2f2; text-align: center; }
        .ttd-container { width: 100%; margin-top: 40px; page-break-inside: avoid; }
        .ttd-kiri { float: left; width: 45%; text-align: center; }
        .ttd-kanan { float: right; width: 45%; text-align: center; }
        .clear { clear: both; }
      </style>
    </head>
    <body>
      <div class="header">
        <h1>${kop[0]}</h1>
        <p>${kop[1]}</p>
      </div>
      <h3>Laporan Rekapitulasi ${jenis}</h3>
      <p class="periode">${textKelas} | Periode: ${tglMulai} s.d ${tglAkhir}</p>
      ${tabelHtml}
      <div class="ttd-container">
        <div class="ttd-kiri">
          <p>Mengetahui,</p>
          <p>Kepala Sekolah</p><br><br><br><br>
          <p><b><u>${kop[4]}</u></b></p>
          <p>NIP. ${kop[5]}</p>
        </div>
        <div class="ttd-kanan">
          <p>................., .........................</p>
          <p>Guru Mata Pelajaran</p><br><br><br><br>
          <p><b><u>${kop[2]}</u></b></p>
          <p>NIP. ${kop[3]}</p>
        </div>
        <div class="clear"></div>
      </div>
    </body>
    </html>
  `;

  let blob = HtmlService.createHtmlOutput(htmlTemplate).getAs('application/pdf').setName(`Laporan_${jenis}_${kelas}.pdf`);
  let folderName = "Rekap My AdTeach";
  let folders = DriveApp.getFoldersByName(folderName);
  let folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
  let file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}
