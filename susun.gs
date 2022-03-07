/*
  *** Konstruktor Data untuk KML Copy Paste Machine ***
  *dibuat oleh Boston Sinaga untuk Faiz Almakmun*

  CARA MENGGUNAKAN:

  -masukan berupa kode Array JSON yang didapat dari "https://bostonsinaga.github.io/"

  -cara input
    -(SINGLE)
      ->  masukkan kode pada "A4"
    -(PARTS / MULTIPLE) (terjadi apabila jumlah karakter kode melampau 50.000 karakter)
      ->  tulis tanda "*" pada "A4" untuk mengisyaratkan input majemuk
          kemudian masukkan potongan kode pada "A5", "A6", "An", ... (berurutan sebanyak jumlah potongan)

  -pastikan seluruh kolom dan baris kosong
  -lalu jalankan script ini...
*/

const sheet = SpreadsheetApp.getActiveSheet();

function susun() {

  const startRow = 4, startCol = 1;
  let currentRow = startRow, currentCol = startCol;
  let data, text, isAvailable = true;

  text = sheet.getRange(startRow, startCol).getValue();

  if (text == '*') { // multiple (
    text = '';
    let rowN = startRow + 1;

    while (true) {
      const cell = sheet.getRange(rowN, startCol);
      if (cell.getValue() != '') {
        text += cell.getValue();
        cell.setHorizontalAlignment('center');
      }
      else {
        break;
      }
      rowN++;
    }

    data = JSON.parse(text);
  }
  else if (text != '') { // single
    data = JSON.parse(text);
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

    sheet.autoResizeRows(1, sheet.getMaxRows());
    sheet.autoResizeColumn(1);
  }
}
