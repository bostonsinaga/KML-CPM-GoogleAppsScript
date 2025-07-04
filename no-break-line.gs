/**
 * Mengubah nilai sel menjadi tanpa baris baru.
 * Blok sel yang ditarget kemudian jalankan kode ini.
 * Blok dari atas ke bawah. JANGAN SEBALIKNYA!
 */
function noBreakLine() {

  let data = [];
  const range = sheet.getActiveRange();
  const curRow = sheet.getActiveCell().getRow();

  for (let i = curRow; i <= range.getNumRows() + curRow - 1; i++) {

    let str = sheet.getRange(i, range.getColumn()).getValue();

    for (let e of str) { 
      if (e == '\n') {
        str = str.replace(e, ' ');
      }
    }

    data.push([str]);
  }

  range.setValues(data);
}

