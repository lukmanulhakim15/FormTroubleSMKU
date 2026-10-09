/*
FORM TROUBLE IT - TANPA TELEGRAM

URUTAN KOLOM:
A NAMA LENGKAP
B EMAIL STAFF
C TANGGAL
D ERROR ISSUE
E TIMESTAMP
F KETERANGAN
G NO ANTRIAN
H REQUEST ID
I STATUS

PENTING:
- Jalankan fungsi rapikanSheetSekarang() SATU KALI setelah memasang script ini.
- Data request yang sudah ada tidak dihapus.
*/

const SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1hacpKFo1J9fnB2-q17YPybrrbftTASsrgR7HcUwPoL8/edit?usp=sharing';
const SHEET_NAME = 'Sheet1';
const NOTIFICATION_EMAIL = '';


/* =========================================================
   GOOGLE SHEET
========================================================= */

function getSheet_() {
  const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);

  if (SHEET_NAME) {
    const sh = ss.getSheetByName(SHEET_NAME);
    if (sh) return sh;
  }

  return ss.getActiveSheet();
}


/* =========================================================
   RAPIIKAN SHEET
========================================================= */

function rapikanSheetSekarang() {
  const sh = getSheet_();

  // Header A-I
  sh.getRange(1, 1, 1, 9).setValues([[
    'NAMA LENGKAP',
    'EMAIL STAFF',
    'TANGGAL',
    'ERROR ISSUE',
    'TIMESTAMP',
    'KETERANGAN',
    'NO ANTRIAN',
    'REQUEST ID',
    'STATUS'
  ]]);

  // Format
  sh.getRange('C:C').setNumberFormat('yyyy-mm-dd');
  sh.getRange('E:E').setNumberFormat('dd/mm/yyyy hh:mm:ss');

  sh.getRange('G:G').setNumberFormat('@');
  sh.getRange('H:H').setNumberFormat('@');
  sh.getRange('I:I').setNumberFormat('@');

  sh.setFrozenRows(1);

  // Lebar kolom
  sh.setColumnWidth(1, 180);
  sh.setColumnWidth(2, 210);
  sh.setColumnWidth(3, 120);
  sh.setColumnWidth(4, 420);
  sh.setColumnWidth(5, 170);
  sh.setColumnWidth(6, 180);
  sh.setColumnWidth(7, 160);
  sh.setColumnWidth(8, 270);
  sh.setColumnWidth(9, 120);

  return 'Selesai: kolom A-I sudah dikembalikan ke urutan yang benar.';
}


/* =========================================================
   GET / TEST WEB APP
========================================================= */

function doGet(e) {

  const action = e && e.parameter
    ? e.parameter.action
    : '';

  // Test koneksi Google Sheet
  if (action === 'test') {

    try {

      const sh = getSheet_();

      return json_({
        result: 'success',
        message: 'Google Sheets terhubung',
        spreadsheet: SpreadsheetApp.getActiveSpreadsheet().getName(),
        sheet: sh.getName()
      });

    } catch (err) {

      return json_({
        result: 'error',
        error: err.message
      });

    }
  }


  // Cek status request
  if (action === 'status') {

    const requestId = e.parameter.requestId;

    if (!requestId) {
      return json_({
        result: 'error',
        error: 'Request ID kosong'
      });
    }

    return json_(findRequest_(requestId));
  }


  return ContentService
    .createTextOutput('FORM TROUBLE IT AKTIF')
    .setMimeType(ContentService.MimeType.TEXT);
}


/* =========================================================
   POST REQUEST TROUBLE IT
========================================================= */

function doPost(e) {

  const lock = LockService.getScriptLock();

  try {

    lock.waitLock(30000);

    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Data form kosong');
    }

    const data = JSON.parse(e.postData.contents);


    // Ambil data
    const nama = clean_(data.nama);
    const email = clean_(data.email);

    const divisi = clean_(
      data.divisi ||
      data.departemen ||
      data.department
    );

    const kategori = clean_(
      data.kategori ||
      data.category
    );

    const deskripsi = clean_(
      data.deskripsi ||
      data.deskripsiKendala ||
      data.errorIssue
    );


    // Validasi
    if (!nama) {
      throw new Error('Nama belum diisi');
    }

    if (!email) {
      throw new Error('Email belum diisi');
    }

    if (!deskripsi) {
      throw new Error('Rincian kendala belum diisi');
    }

    if (!isValidEmail_(email)) {
      throw new Error(
        'Format email tidak valid: ' + email
      );
    }


    const sh = getSheet_();


    // Pastikan header tetap A-I
    sh.getRange(1, 1, 1, 9).setValues([[
      'NAMA LENGKAP',
      'EMAIL STAFF',
      'TANGGAL',
      'ERROR ISSUE',
      'TIMESTAMP',
      'KETERANGAN',
      'NO ANTRIAN',
      'REQUEST ID',
      'STATUS'
    ]]);


    // Data request
    const now = new Date();

    const nomorAntrian = buatNomorAntrian_();

    const requestId =
      clean_(data.requestId) ||
      Utilities.getUuid();


    const errorIssue =
      kategori
        ? kategori + ' - ' + deskripsi
        : deskripsi;


    // Simpan ke Spreadsheet
    sh.appendRow([
      nama,           // A
      email,          // B
      now,            // C
      errorIssue,     // D
      now,            // E
      divisi,         // F
      nomorAntrian,   // G
      requestId,      // H
      'Menunggu'      // I
    ]);


    // Cache status
    CacheService
      .getScriptCache()
      .put(
        'request_' + requestId,
        JSON.stringify({
          queueNumber: nomorAntrian,
          status: 'Menunggu'
        }),
        21600
      );


    // Email staff
    try {

      sendStaffEmail_(
        email,
        nama,
        nomorAntrian,
        requestId,
        divisi,
        kategori,
        deskripsi
      );

    } catch (err) {

      console.error(
        'Email staff gagal: ' +
        err.message
      );

    }


    // Email admin jika diaktifkan
    if (NOTIFICATION_EMAIL) {

      try {

        sendAdminEmail_(
          NOTIFICATION_EMAIL,
          nama,
          email,
          nomorAntrian,
          requestId,
          divisi,
          kategori,
          deskripsi
        );

      } catch (err) {

        console.error(
          'Email admin gagal: ' +
          err.message
        );

      }
    }


    // Response
    return json_({

      result: 'success',

      requestId: requestId,

      queueNumber: nomorAntrian,

      status: 'Menunggu'

    });


  } catch (err) {

    return json_({

      result: 'error',

      error: err.message

    });


  } finally {

    try {
      lock.releaseLock();
    } catch (_) {}

  }
}


/* =========================================================
   NOMOR ANTRIAN
========================================================= */

function buatNomorAntrian_() {

  const props =
    PropertiesService.getScriptProperties();

  const tz =
    Session.getScriptTimeZone() ||
    'Asia/Jakarta';

  const tanggal =
    Utilities.formatDate(
      new Date(),
      tz,
      'yyyyMMdd'
    );

  const key =
    'ANTRIAN_' + tanggal;


  let nomor =
    Number(
      props.getProperty(key) || 0
    ) + 1;


  props.setProperty(
    key,
    String(nomor)
  );


  return (
    'IT-' +
    tanggal +
    '-' +
    String(nomor).padStart(3, '0')
  );
}


/* =========================================================
   EMAIL STAFF
========================================================= */

function sendStaffEmail_(
  email,
  nama,
  nomor,
  requestId,
  divisi,
  kategori,
  deskripsi
) {

  GmailApp.sendEmail(

    email,

    'Request IT berhasil diterima - ' +
    nomor,

    'Request IT diterima. No. Antrian: ' +
    nomor +
    '. Status: Menunggu.',

    {

      name: 'IT Support',

      htmlBody:

        '<h2>Request IT berhasil diterima</h2>' +

        '<p>Halo <b>' +
        escapeHtml_(nama) +
        '</b>,</p>' +

        '<p>Request Trouble IT kamu sudah berhasil diterima.</p>' +

        '<h2>NO ANTRIANMU<br>' +
        escapeHtml_(nomor) +
        '</h2>' +

        '<p><b>Status:</b> Menunggu</p>' +

        '<p><b>Departement:</b> ' +
        escapeHtml_(divisi || '-') +
        '</p>' +

        '<p><b>Kategori:</b> ' +
        escapeHtml_(kategori || '-') +
        '</p>' +

        '<p><b>Kendala:</b><br>' +
        escapeHtml_(deskripsi)
          .replace(/\n/g, '<br>') +
        '</p>' +

        '<p><b>Request ID:</b> ' +
        escapeHtml_(requestId) +
        '</p>'

    }

  );
}


/* =========================================================
   EMAIL ADMIN
========================================================= */

function sendAdminEmail_(
  to,
  nama,
  email,
  nomor,
  requestId,
  divisi,
  kategori,
  deskripsi
) {

  GmailApp.sendEmail(

    to,

    'Request IT Baru - ' +
    nomor,

    'Request IT baru: ' +
    nomor,

    {

      name: 'IT Support',

      htmlBody:

        '<h2>Request Trouble IT Baru</h2>' +

        '<p><b>No. Antrian:</b> ' +
        escapeHtml_(nomor) +
        '</p>' +

        '<p><b>Nama:</b> ' +
        escapeHtml_(nama) +
        '</p>' +

        '<p><b>Email:</b> ' +
        escapeHtml_(email) +
        '</p>' +

        '<p><b>Departement:</b> ' +
        escapeHtml_(divisi || '-') +
        '</p>' +

        '<p><b>Kategori:</b> ' +
        escapeHtml_(kategori || '-') +
        '</p>' +

        '<p><b>Kendala:</b><br>' +
        escapeHtml_(deskripsi)
          .replace(/\n/g, '<br>') +
        '</p>' +

        '<p><b>Request ID:</b> ' +
        escapeHtml_(requestId) +
        '</p>'

    }

  );
}


/* =========================================================
   TEST EMAIL
========================================================= */

function authorizeAndTestEmail() {

  const to =
    Session.getEffectiveUser().getEmail();


  GmailApp.sendEmail(

    to,

    'TEST - Email Trouble IT',

    'Tes email berhasil.',

    {

      name: 'IT Support',

      htmlBody:
        '<h2>TEST - Email Trouble IT</h2>' +
        '<p>Apps Script memiliki izin email.</p>'

    }

  );
}


/* =========================================================
   CEK STATUS REQUEST
========================================================= */

function findRequest_(requestId) {

  const cache =
    CacheService
      .getScriptCache()
      .get('request_' + requestId);


  if (cache) {

    const d =
      JSON.parse(cache);


    return {

      result: 'success',

      found: true,

      requestId: requestId,

      queueNumber: d.queueNumber,

      status: d.status

    };

  }


  const values =
    getSheet_()
      .getDataRange()
      .getValues();


  for (
    let i = values.length - 1;
    i >= 1;
    i--
  ) {

    if (
      String(values[i][7]) ===
      String(requestId)
    ) {

      return {

        result: 'success',

        found: true,

        requestId: requestId,

        queueNumber: values[i][6],

        status:
          values[i][8] ||
          'Menunggu'

      };

    }

  }


  return {

    result: 'success',

    found: false

  };
}


/* =========================================================
   HELPER
========================================================= */

function clean_(v) {

  return v == null
    ? ''
    : String(v).trim();

}


function isValidEmail_(v) {

  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/
    .test(v);

}


function escapeHtml_(v) {

  return String(
    v == null ? '' : v
  )

    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');

}


function json_(data) {

  return ContentService

    .createTextOutput(
      JSON.stringify(data)
    )

    .setMimeType(
      ContentService.MimeType.JSON
    );

}

