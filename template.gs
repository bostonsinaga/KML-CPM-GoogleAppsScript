/*
  *** Membuat template kosong ***
  -judul tabel
  -style dasar
  (JIKA TIDAK BEKERJA, COBA HAPUS "*" DI AD:6000)
*/

function keepMaximumCells() {

  let max = sheet.getMaxColumns();

  if (max < 30) {
    sheet.insertColumns(max, 30 - max);
    sheet.getRangeList(['A1:AD1']).setBackground('white');
  }
  else if (max > 30) {
    sheet.deleteColumns(31, max - 30);
  }

  max = sheet.getMaxRows();

  if (max < 6000) {
    sheet.insertRows(max, 6000 - max);
  }
  else if (max > 6000) {
    sheet.deleteRows(6001, max - 6000);
  }
}

function template() {

  if (sheet.getMaxColumns() != 30 || sheet.getMaxRows() != 6000) {
    keepMaximumCells();
  }

  if (sheet.getRange('AD6000').getValue() != '*') {
    sheet.getRange('AD6000').setValue('*');

    sheet.getRangeList(['A:AD'])
      .setFontFamily('Arial')
      .setFontSize(10)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('center');

    sheet.getRangeList(['A1:G3', 'J1:N3', 'Q1:V3', 'Y1:AD3'])
      .setBackground('#DDDDDD')
      .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

    // judul
    let rangeCode = ['A1:G1', 'J1:N1', 'Q1:V1', 'Y1:AD1'];
    let judul = ['JALUR', 'CLOSURE', 'ODP', 'TIANG'];

    // sub judul (ROW 2)
    let subJudul = [
      [
        ['A2', 'B2', 'C2', 'D2', 'E2', 'F2', 'G2'],
        ['J2', 'K2', 'L2', 'M2', 'N2'],
        ['Q2', 'R2', 'S2', 'T2', 'U2', 'V2'],
        ['Y2', 'Z2', 'AA2', 'AB2', 'AC2', 'AD2']
      ],
      [
        ['NAMA', 'KETERANGAN', 'HARGA/METER (RP)', 'TOTAL JARAK (M)', 'TOTAL HARGA (RP)', 'TIKOR AWAL', 'TIKOR AKHIR'],
        ['NAMA', 'LATITUDE', 'LONGITUDE', 'HARGA (RP)', 'KETERANGAN'],
        ['NAMA', 'LATITUDE', 'LONGITUDE', 'KAPASITAS', 'HARGA (RP)', 'KETERANGAN'],
        ['NAMA', 'LATITUDE', 'LONGITUDE', 'TINGGI (M)', 'HARGA (RP)', 'KETERANGAN']
      ]
    ];

    // formulas
    let subJudulFormula = [
      [
        ['A3', '=COUNTA(B4:B6000)'],
        ['B3:C3', 'merge'],
        ['D3', '=SUM(D4:D6000)'],
        ['E3', '=SUM(E4:E6000)'],
        ['F3:G3', 'merge']
      ],
      [
        ['J3', '=COUNTA(J4:J6000)'],
        ['K3:L3', 'merge'],
        ['M3', '=SUM(M4:M6000)'],
      ],
      [
        ['Q3', '=COUNTA(Q4:Q6000)'],
        ['R3:T3', 'merge'],
        ['U3', '=SUM(U4:U6000)'],
      ],
      [
        ['Y3', '=COUNTA(Y4:Y6000)'],
        ['Z3:AB3', 'merge'],
        ['AC3', '=SUM(AC4:AC6000)'],
      ]
    ];

    for (let i = 0; i < 4; i++) {

      sheet.getRange(rangeCode[i]).merge();
      sheet.getRange(rangeCode[i].slice(0, 2)).setValue(judul[i]);

      // sub judul
      let ct = subJudul[0][i].length;
      for (let j = 0; j < ct; j++) {
        sheet.getRange(subJudul[0][i][j]).setValue(subJudul[1][i][j]);
      }

      ct = subJudulFormula[i].length;
      for (let j = 0; j < ct; j++) {
        if (subJudulFormula[i][j][1] == 'merge') {
          sheet.getRange(subJudulFormula[i][j][0]).merge();
        }
        else {
          sheet.getRange(subJudulFormula[i][j][0]).setFormula(subJudulFormula[i][j][1]);
        }
      }
    }

    sheet.setRowHeights(1, sheet.getMaxRows(), 21);
    sheet.setColumnWidths(1, sheet.getMaxColumns(), 100);
    sheet.setFrozenRows(3);
  }
}
