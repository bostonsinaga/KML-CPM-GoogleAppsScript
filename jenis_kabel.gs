// JENIS KABEL BERDASARKAN PANJANGNYA
// **block dari atas ke bawah

function jenisKabel() {

  let data = [], nameData = [];
  const range = sheet.getActiveRange();
  const curCell = sheet.getActiveCell(); 
  const curRow = curCell.getRow();

  const RANGE_ROWS = range.getNumRows();

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
    
    let preData = '';
    if (i != curRow) {
      preData = sheet.getRange(i, 2).getValue();
      if (preData != '') preData = '\n' + preData;
      else preData = '';
    }

    let udaraTanam = 'UDARA';
    if (isTanam) udaraTanam = 'TANAM';

    const getCore = (def) => {
      const core = sheet.getRange(i, 3).getValue();
      if (core == '') return [def, def / 6];
      return [core, core / 6];
    }

    const setData = (jenis, core, tube) => {
      data.push([`${jenis} ${udaraTanam} XNODE ${core}C/${tube}T ${preData}`]);
    };

    if (panjang > 10000) {
      const core  = getCore(48);
      setData('BACKBONE', core[0], core[1]);
    }
    else if (panjang > 750) {
      const core  = getCore(24);
      setData('BACKBONE', core[0], core[1]);
    }
    else {
      const core  = getCore(12);
      setData('AKSES', core[0], core[1]);
    }
  }

  range.setValues(data);
  sheet.getRange(curRow, 1, RANGE_ROWS, 1).setValues(nameData);
}

