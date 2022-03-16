// TIDAK DIREKOMENDASIKAN.. (Karena agak rumit)
// MENAMBAHKAN DESKRIPSI DENGAN TANGGAL
// -- CLICK PADA CELL TARGET (lebih tepatnya di 'Keterangan')
// -- ATUR NILAI VARIABLE 'AKHIR_DATA_INDEX'

function isiKosong() {

  const bulanStr = [
    ['Januari', 31], ['Februari', 28], ['Maret', 31], ['April', 30],
    ['Mei', 31],['Juni', 30], ['Juli', 31], ['Agustus', 31],
    ['September', 30], ['Oktober', 31], ['November', 30], ['Desember', 31],
  ];





 ///////////// YANG PERLU DIEDIT //////////////

 const AKHIR_DATA_INDEX = 98; // Masukkan index terakhir kolom target (nama)

 ////////////////////////////////////////////////////////
 




  let JUMLAH_ROW;
  const JUMLAH_TOTAL = AKHIR_DATA_INDEX - 3;

  let target;

  for (let BULAN_INDEX = 1; BULAN_INDEX <= 12; BULAN_INDEX++) {

    let TANGGAL_AWAL;
    let TANGGAL_AKHIR;
    
    const BULAN_TAHUN = ` ${bulanStr[BULAN_INDEX - 1][0]} 2020`;

    if (BULAN_INDEX == 12) {
      JUMLAH_ROW = JUMLAH_TOTAL - Math.floor(JUMLAH_TOTAL / 12) * 11;
      TANGGAL_AKHIR = 20;
    }
    else {
      TANGGAL_AKHIR = bulanStr[BULAN_INDEX - 1][1];
      JUMLAH_ROW = Math.floor(JUMLAH_TOTAL / 12);
    }

    if (BULAN_INDEX == 1) {
      TANGGAL_AWAL = 8;
    }
    else {
      TANGGAL_AWAL = 1;
    }

    const deskripsi = [];

    if (BULAN_INDEX == 1) {
      target = sheet.getRange(
        sheet.getActiveCell().getRow(), 
        sheet.getActiveCell().getColumn(),
        JUMLAH_ROW,
        1
      );
    }
    else {
      target = sheet.getRange(
        (BULAN_INDEX == 12 ? Math.floor(JUMLAH_TOTAL / 12) : JUMLAH_ROW)
          * (BULAN_INDEX - 1) + 4,
        target.getColumn(),
        JUMLAH_ROW,
        1
      );
    }

    for (let i = 1; i <= JUMLAH_ROW; i++) {
      let tanggal = Math.ceil(TANGGAL_AKHIR / JUMLAH_ROW * i);
      if (tanggal < TANGGAL_AWAL) tanggal = TANGGAL_AWAL;
      else if (tanggal > TANGGAL_AKHIR) tanggal = TANGGAL_AKHIR;
      deskripsi.push([target.getValues()[i-1][0] + ' ' + tanggal + BULAN_TAHUN]);
    }

    target.setValues(deskripsi);
  }
}
