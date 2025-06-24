/**
 * !!! PERHATIAN... JIKA BERALIH KE SHEET LAIN, MOHON UNTUK MERELOAD LAMAN INI !!!
 * Karena sheet target dapat tidak terdeteksi dan akan mengakibatkan target eksekusi
 * menjadi sheet sebelumnya. Berlaku untuk semua kode sumber di repo ini.
 * 
 * Konstruktor Data untuk KML Copy Paste Machine.
 * Dibuat oleh Boston Sinaga untuk Faiz Almakmun.
 * 
 * CARA MENGGUNAKAN:
 * - Masukan berupa kode Array JSON yang didapat dari "https://github.com/bostonsinaga/KML-CPM-WebPage".
 * - Cara input
 *   - SINGLE: masukkan kode pada "A4".
 *   - PARTS / MULTIPLE:
 *     Terjadi apabila jumlah karakter kode melampau 50.000 karakter.
 *     Tulis tanda "*" pada "A4" untuk mengisyaratkan input majemuk.
 *     Kemudian masukkan potongan kode pada "A5", "A6", "An", ...
 *     Ini berurutan sebanyak jumlah potongan.
 * - Untuk melanjutkan data silahkan masukkan data pada kolom 'A' di baris lanjutan.
 *   Kemudian isi nilai cell 'AC6000' dengan urutan baris tersebut.
 *   Pastikan semua kolom pada baris itu kosong, karena data yang ada akan ditimpa.
 * - Pastikan seluruh kolom dan baris kosong.
 * - Lalu jalankan script ini.
 */

const sheet = SpreadsheetApp.getActiveSheet();

function susun() {
  const AC6000 = sheet.getRange('AC6000').getValue(); // 'startRow' adjust
  const startRow = AC6000 == '' ? 4 : AC6000;
  sheet.getRange('AC6000').clear();
  const startCol = 1;

  let currentRow = startRow, currentCol = startCol;
  let data, text, isAvailable = true;

  const firstRange = sheet.getRange(startRow, startCol);

  const firstRangeDefault = () => {
    firstRange
      .clear()
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  }

  text = firstRange.getValue();

  // multiple
  if (text == '*') {
    text = '';
    firstRangeDefault();
    let rowN = startRow + 1;

    while (true) {
      const cell = sheet.getRange(rowN, startCol);
      if (cell.getValue() != '') {
        text += cell.getValue();
        cell
          .clear()
          .setHorizontalAlignment('center')
          .setVerticalAlignment('middle');
      }
      else {
        break;
      }
      rowN++;
    }

    data = JSON.parse(text);
  }
  // single
  else if (text != '') {
    data = JSON.parse(text);
    firstRangeDefault();
  }
  else {
    isAvailable = false;
    console.log('Data tidak ditemukan...');
  }

  if (isAvailable) {
    template();

    // cetak data
    for (let i = 0; i < data.length; i++) {
      if (data[i] == '=>') {
        currentCol++;
        currentRow = startRow;
      } else {
        sheet.getRange(currentRow, currentCol).setValue(data[i]);
        currentRow++;
      }
    }

    resize();
  }
}

