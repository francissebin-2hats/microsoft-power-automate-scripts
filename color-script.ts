function main(
  workbook: ExcelScript.Workbook,
  confirmedCol: number = 5,// Column E
  etdCol: number = 17) // Column Q
{

  const sheet = workbook.getActiveWorksheet();

  // Get used rows
  const usedRange = sheet.getUsedRange();
  const rowCount = usedRange.getRowCount();

  for (let i = 1; i < rowCount; i++) { // starting from row 2
      const row = usedRange.getCell(i, 0).getEntireRow();

      const confirmedDate = usedRange.getCell(i, confirmedCol - 1).getValue() as number;
      const etdDate = usedRange.getCell(i, etdCol - 1).getValue() as number;

      // Skip if dates are missing
      if (!confirmedDate || !etdDate) {
          row.getFormat().getFill().clear();
          continue;
      }

      // Calculate date difference
      const diff = etdDate - confirmedDate;

      console.log(confirmedDate + ' ' + etdDate + ' ' + diff);
      // Apply colors
      if (diff <= 0) {
          // console.log('no color');
          // No color
          row.getFormat().getFill().clear();
      }
      else if (diff >= 1 && diff <= 9) {
          // console.log('yellow color');
          // Yellow
          row.getFormat().getFill().setColor("Yellow");
      }
      else if (diff >= 10) {
          // console.log('red color');
          // Red
          row.getFormat().getFill().setColor("Red");
      }
  }
}