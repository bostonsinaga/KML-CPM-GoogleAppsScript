// MENAMBAHKAN DESKRIPSI DENGAN TANGGAL
// -- CLICK PADA CELL TARGET (lebih tepatnya di 'Keterangan')
// -- ATUR NILAI VARIABLE YANG PERLU DIEDIT (beberapa variable jika diganti mungkin dapat membuat program error)
// -- JIKA BULAN LEBIH DARI 1 DAN SAMPAI KE DESEMBER
//    ISI DULU SAMPAI NOVEMBER BARU KE DESEMBER (TOGGLE ADA DI BAWAH)

function isiKosong() {

  let BULAN_AWAL, BULAN_AKHIR;
  function SET_BULAN_AWAL_AKHIR(awal, akhir) {
    if (awal == undefined && akhir == undefined) {
      BULAN_AWAL = 12;
      BULAN_AKHIR = 12;
    }
    else {
      BULAN_AWAL = awal;
      BULAN_AKHIR = akhir;
    }
  }






 ///////////// YANG PERLU DIEDIT //////////////

 const AKHIR_DATA_INDEX = undefined; // Masukkan index terakhir kolom target (nama)
 const TAHUN = undefined;
 const PENGHUBUNG = '';

 SET_BULAN_AWAL_AKHIR( // contoh "1, 11" (januari-november) --V
    // 1, 11
 ); // ^-- masukan kosong jika hanya bulan 12

 ////////////////////////////////////////////////////////
 





  let JUMLAH_ROW;
  const JUMLAH_TOTAL = AKHIR_DATA_INDEX - sheet.getCurrentCell().getRow() + 1;

  let target, counter = 0;

  for (let BULAN_INDEX = BULAN_AWAL; BULAN_INDEX <= BULAN_AKHIR; BULAN_INDEX++) {

    let TANGGAL_AWAL;
    let TANGGAL_AKHIR;
    
    const BULAN_TAHUN = ` ${bulanStr[BULAN_INDEX - 1][0]} ${TAHUN}`;

    const totalBulan = BULAN_AKHIR - BULAN_AWAL + 1;
    if (BULAN_INDEX == 12) {
      JUMLAH_ROW = JUMLAH_TOTAL - Math.floor(JUMLAH_TOTAL / totalBulan) * (totalBulan - 1);
      TANGGAL_AKHIR = 20;
    }
    else {
      TANGGAL_AKHIR = bulanStr[BULAN_INDEX - 1][1];
      const jlh = parseFloat(JUMLAH_TOTAL / (BULAN_AKHIR - BULAN_AWAL + 1));
      if (jlh % 1 * 10 >= 5) JUMLAH_ROW = Math.ceil(jlh);
      else JUMLAH_ROW = Math.floor(jlh);
    }

    if (BULAN_INDEX == 1) {
      TANGGAL_AWAL = 8;
    }
    else {
      TANGGAL_AWAL = 1;
    }

    const deskripsi = [];

    if (BULAN_INDEX == BULAN_AWAL) {
      target = sheet.getRange(
        sheet.getActiveCell().getRow(), 
        sheet.getActiveCell().getColumn(),
        JUMLAH_ROW,
        1
      );
    }
    else {
      target = sheet.getRange(
        (BULAN_INDEX == 12 ? Math.floor(JUMLAH_TOTAL / 12) : JUMLAH_ROW) * counter + 4,
        target.getColumn(),
        JUMLAH_ROW,
        1
      );
    }

    for (let i = 1; i <= JUMLAH_ROW; i++) {
      let tanggal = Math.ceil(TANGGAL_AKHIR / JUMLAH_ROW * i);
      if (tanggal < TANGGAL_AWAL) tanggal = TANGGAL_AWAL;
      else if (tanggal > TANGGAL_AKHIR) tanggal = TANGGAL_AKHIR;
      deskripsi.push([target.getValues()[i-1][0] + PENGHUBUNG + tanggal + BULAN_TAHUN]);
    }

    target.setValues(deskripsi);
    counter++;
  }
}
