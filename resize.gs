/*
  *** Menyesuaikan ukuran cell dengan value-nya ***
*/

function resize() {
  sheet.autoResizeRows(1, sheet.getMaxRows());
  for (let i = 1; i <= sheet.getMaxColumns(); i++) {
    sheet.autoResizeColumn(i);
  }
}
