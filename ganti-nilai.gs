/**
 * Ganti data spesifik pada kumpulan sel.
 * Blok dari atas ke bawah.
 */

function gantiNilai() {
  /////////////////////////////////////////** YANG PERLU DIGANTI *////////////////////////////////

  // karakter awal sampai akhir (inklusif)
  const TITIK = [
    'L',
    ''
  ];

  // Setiap kata sama. Beda hanya 'uppercase' atau 'lowercase'
  const PRASA_UMUM = [
    'atitude'
  ];

  const NILAI_BARU = [
    ''
  ];

  /**
   * Contoh bawaan:
   * (AWAL, AKHIR, VALUE) => false
   */
  const KONDISI_KHUSUS = [
    (AWAL, AKHIR, VALUE) => false
  ];

  ////////////////////////////////////////////////////////////////////////////////////////////////

  let data = [];
  const range = sheet.getActiveRange();
  const curCell = sheet.getActiveCell();

  for (let counter = curCell.getRow();
    counter <= range.getNumRows() + curCell.getRow() - 1;
    counter++
  ) {
    let curVal = sheet.getRange(
      counter, curCell.getColumn()
    ).getValue();

    for (let i = 0; i < curVal.length; i++) {
      let str = curVal[i];

      if (str == TITIK[0]) {
        str = '';

        let j;
        for (j = 1; j <= PRASA_UMUM[0].length; j++) {
          str += curVal[i+j];
        }
        j--;

        let isAvailable = false;
        for (e of PRASA_UMUM) {
          if (str == e) {
            isAvailable = true;
            break;
          }
        }

        if (isAvailable) {
          let k = 2;
          const getLastChr = () => i+j+k;

          while(
            curVal[getLastChr()] != TITIK[1] &&
            getLastChr() < curVal.length - 1
          ) {
            k++;
          }

          let ct = 0;
          for (e of KONDISI_KHUSUS) {
            if (e(i, getLastChr(), curVal)) {

              data.push([curVal.replace(
                curVal.slice(i, getLastChr()+1), NILAI_BARU[ct]
              )]);
              ct = -1;
              break;
            }
            ct++;
          }

          if (ct != -1) {
            data.push([curVal.replace(
              curVal.slice(i, getLastChr()+1), NILAI_BARU[NILAI_BARU.length - 1]
            )]);
          }

          break;
        }
      }

      if (i >= curVal.length - 1) {
        data.push([curVal]);
      }
    }

    if (curVal.length == 0) {
      data.push([curVal]);
    }
  }

  range.setValues(data);
}

