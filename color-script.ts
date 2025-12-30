function main(
  workbook: ExcelScript.Workbook,
  confirmedCol: number = 5, // Column E
  etdCol: number = 17
) {
  // Column Q
  const sheet = workbook.getActiveWorksheet();

  // Get used rows
  const usedRange = sheet.getUsedRange();
  const rowCount = usedRange.getRowCount();

  for (let i = 1; i < rowCount; i++) {
    // starting from row 2
    const confirmedDate = usedRange
      .getCell(i, confirmedCol - 1)
      .getValue() as number;
    const etdDate = usedRange.getCell(i, etdCol - 1).getValue() as number;

    // Determine color based on date difference
    let fillColor: string | null = null;
    if (confirmedDate && etdDate) {
      // Calculate date difference
      const diff = etdDate - confirmedDate;
      // Apply colors
      if (diff >= 1 && diff <= 9) {
        // Yellow
        fillColor = "Yellow";
      } else if (diff >= 10) {
        // Red
        fillColor = "Red";
      }
      // If diff <= 0, fillColor remains null (no color)
    }

    // Color cells A-U (columns 1-21, indices 0-20)
    for (let col = 0; col < 21; col++) {
      const cell = usedRange.getCell(i, col);
      if (fillColor) {
        cell.getFormat().getFill().setColor(fillColor);
      } else {
        cell.getFormat().getFill().clear();
      }
    }
  }
}
