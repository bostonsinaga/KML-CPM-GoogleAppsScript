// JENIS KABEL BERDASARKAN PANJANGNYA
// **block dari atas ke bawah
// **masukkan data tanggal di cell awal dalam format "[tanggal-awal, bulan-awal, tanggal-akhir, bulan-akhir, tahun]"

function noBreakLine() {

  let data = [], nameData = [];
  const range = sheet.getActiveRange();
  const curCell = sheet.getActiveCell(); 
  const curRow = curCell.getRow();

  const ARR_WAKTU = JSON.parse(curCell.getValue());
  
  let JUMLAH_HARI = 0;

  ARR_WAKTU[1] - ARR_WAKTU[3] + 1; 

  bulanStr.forEach((e, i) => {
    if (i >= ARR_WAKTU[1] - 1 && i <= ARR_WAKTU[3] - 1) {
      JUMLAH_HARI += e[1];
      if (i == ARR_WAKTU[1]) JUMLAH_HARI -= ARR_WAKTU[0] - 1;
      else if (i == ARR_WAKTU[3]) JUMLAH_HARI -= e[1] - ARR_WAKTU[2];
    }
  });

  const RANGE_ROWS = range.getNumRows();
  const HARI_RATE = parseFloat(JUMLAH_HARI / RANGE_ROWS);
  let WAKTU;

  let HARI = ARR_WAKTU[0], BULAN = ARR_WAKTU[1] - 1;

  for (let i = curRow; i <= RANGE_ROWS + curRow - 1; i++) {

    let isTanam = false;
    const nama = sheet.getRange(i, 1).getValue();
    if (nama.slice(nama.length - 2) == '@-') {
      nameData.push([nama.slice(0, nama.length - 2)]);
      isTanam = true;
    }
    else {
      nameData.push([nama]);
    }

    const panjang = parseInt(sheet.getRange(i, 4).getValue());

    if (HARI % 1 * 10 < 5) HARI = Math.floor(HARI);
    else HARI = Math.ceil(HARI);

    WAKTU = `${HARI} ${bulanStr[BULAN][0]} ${ARR_WAKTU[4]}`;
    
    let preData = '';
    if (i != curRow) {
      preData = sheet.getRange(i, 2).getValue();
      if (preData != '') preData = '\n' + preData;
      else preData = '';
    }

    let udaraTanam = 'UDARA';
    if (isTanam) udaraTanam = 'TANAM';

    if (panjang > 750) data.push([`BACKBONE ${udaraTanam} XNODE 24C/4T ${WAKTU}${preData}`]);
    else data.push([`AKSES ${udaraTanam} XNODE 12C/2T ${WAKTU}${preData}`]);

    HARI += HARI_RATE;
    if (HARI > bulanStr[BULAN][1]) {
      HARI = 1;
      BULAN++;
      if (BULAN > ARR_WAKTU[3]) BULAN = ARR_WAKTU[3];
    }
  }

  range.setValues(data);
  sheet.getRange(curRow, 1, RANGE_ROWS, 1).setValues(nameData);
}

