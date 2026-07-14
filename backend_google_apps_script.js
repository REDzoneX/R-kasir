/**
 * R-kasir Apps Script Integration
 * Integrasi Cloud Otomatis untuk mencatat Transaksi Toko, Inventori & Log Kasir secara real-time.
 */

function doGet(e) {
  // Meluncurkan halaman HTML web kasir
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('R-kasir - Aplikasi Kasir Toko Profesional')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Membuat tabel penampung otomatis jika file sheets Anda masih kosong
 */
function createSheetsSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Buat Sheet Transaksi Penjualan
  let trxSheet = ss.getSheetByName('Transaksi_Penjualan');
  if (!trxSheet) {
    trxSheet = ss.insertSheet('Transaksi_Penjualan');
    trxSheet.appendRow(['ID TRANSAKSI', 'TANGGAL & WAKTU', 'NAMA PETUGAS KASIR', 'DAFTAR PRODUK (QTY)', 'TOTAL BELANJA (Rp)', 'CASH BAYAR (Rp)', 'KEMBALIAN (Rp)']);
    trxSheet.getRange('A1:G1').setBackground('#312e81').setFontColor('#ffffff').setFontWeight('bold');
  }

  // 2. Buat Sheet Log Aktivitas Kasir
  let logSheet = ss.getSheetByName('Log_Aktivitas');
  if (!logSheet) {
    logSheet = ss.insertSheet('Log_Aktivitas');
    logSheet.appendRow(['TANGGAL', 'WAKTU', 'PETUGAS / PENGGUNA', 'JENIS AKTIVITAS', 'DETAIL CATATAN PENGGUNA']);
    logSheet.getRange('A1:E1').setBackground('#1e1b4b').setFontColor('#ffffff').setFontWeight('bold');
  }
}

/**
 * Mencatat log aktivitas setiap kasir login, input stok baru, maupun aktivitas lainnya
 */
function writeLogToSheet(logItem) {
  try {
    createSheetsSetup(); // Pastikan lembar log sudah ada
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Log_Aktivitas');
    
    // Simpan baris data log baru
    sheet.appendRow([
      logItem.date,
      logItem.time,
      logItem.user,
      logItem.type.toUpperCase(),
      logItem.text
    ]);
    
    return { success: true, message: 'Log aktivitas berhasil disinkronisasi ke Google Sheet.' };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Mencatat detail invoice transaksi penjualan kasir
 */
function writeTransactionToSheet(trx) {
  try {
    createSheetsSetup(); // Pastikan lembar transaksi sudah ada
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Transaksi_Penjualan');
    
    // Simpan baris data transaksi belanja baru
    sheet.appendRow([
      trx.id,
      trx.date,
      trx.cashier,
      trx.items,
      trx.totalPrice,
      trx.cashPaid,
      trx.cashChange
    ]);
    
    return { success: true, message: 'Data transaksi penjualan berhasil dicatat di cloud Google Sheet.' };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}