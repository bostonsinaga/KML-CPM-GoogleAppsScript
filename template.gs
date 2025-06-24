/**
 * Membuat template kosong untuk judul tabel dan style dasar.
 * Jika tidak bekerja coba hapus '*' DI AK:6000.
 */
function keepMaximumCells() {
  let max = sheet.getMaxColumns();

  if (max < 30) {
    sheet.insertColumns(max, 30 - max);
    sheet.getRangeList(['A1:AK1']).setBackground('white');
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

  if (sheet.getRange('AK6000').getValue() != '*') {
    sheet.getRange('AK6000').setValue('*');

    sheet.getRangeList(['A:AK'])
      .setFontFamily('Arial')
      .setFontSize(10)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    sheet.getRangeList(['A1:G3', 'J1:N3', 'Q1:U3', 'X1:AC3', 'AF1:AK3'])
      .setFontWeight('bold')
      .setBackground('#DDDDDD')
      .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

    // judul
    let rangeCode = ['A1:G1', 'J1:N1', 'Q1:U1', 'X1:AC1', 'AF1:AK1'];
    let judul = ['JALUR', 'CLOSURE', 'HAND HOLE', 'ODP', 'TIANG'];

    // sub judul (ROW 2)
    let subJudul = [
      [
        ['A2', 'B2', 'C2', 'D2', 'E2', 'F2', 'G2'],
        ['J2', 'K2', 'L2', 'M2', 'N2'],
        ['Q2', 'R2', 'S2', 'T2', 'U2'],
        ['X2', 'Y2', 'Z2', 'AA2', 'AB2', 'AC2'],
        ['AF2', 'AG2', 'AH2', 'AI2', 'AJ2', 'AK2']
      ],
      [
        ['NAMA', 'KETERANGAN', 'HARGA/METER (RP)', 'TOTAL JARAK (M)', 'TOTAL HARGA (RP)', 'TIKOR AWAL', 'TIKOR AKHIR'],
        ['NAMA', 'LATITUDE', 'LONGITUDE', 'HARGA (RP)', 'KETERANGAN'],
        ['NAMA', 'LATITUDE', 'LONGITUDE', 'HARGA (RP)', 'KETERANGAN'],
        ['NAMA', 'LATITUDE', 'LONGITUDE', 'KAPASITAS', 'HARGA (RP)', 'KETERANGAN'],
        ['NAMA', 'LATITUDE', 'LONGITUDE', 'TINGGI (M)', 'HARGA (RP)', 'KETERANGAN']
      ]
    ];

    // formulas
    let subJudulFormula = [
      [
        ['A3', '=COUNTA(A4:A6000)'],
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
        ['R3:S3', 'merge'],
        ['T3', '=SUM(T4:T6000)'],
      ],
      [
        ['X3', '=COUNTA(X4:X6000)'],
        ['Y3:AA3', 'merge'],
        ['AB3', '=SUM(AB4:AB6000)'],
      ],
      [
        ['AF3', '=COUNTA(AF4:AF6000)'],
        ['AG3:AI3', 'merge'],
        ['AJ3', '=SUM(AJ4:AJ6000)'],
      ]
    ];

    for (let i = 0; i < 5; i++) {

      sheet.getRange(rangeCode[i]).merge();

      let colonIndex;
      if (rangeCode[i].length > 6) colonIndex = 3;
      else colonIndex = 2;

      sheet.getRange(rangeCode[i].slice(0, colonIndex)).setValue(judul[i]);

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

    // set number format
    sheet.getRange('C:E').setNumberFormat('#,##0');
    sheet.getRange('M:M').setNumberFormat('#,##0');
    sheet.getRange('T:T').setNumberFormat('#,##0');
    sheet.getRange('AA:AB').setNumberFormat('#,##0');
    sheet.getRange('AI:AJ').setNumberFormat('#,##0');

    resize();
    sheet.setFrozenRows(3);
    sheet.getRange('H1').setValue('\n');
  }
}

