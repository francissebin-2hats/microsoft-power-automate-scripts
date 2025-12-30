function main(
  workbook: ExcelScript.Workbook,
  columnLetter: string = "E",
  currentPassword: string = "GreenMoon",
  lock: boolean = true
) {
  // const startTime = Date.now();
  // console.log("Script execution started");
  let sheet = workbook.getActiveWorksheet();

  // Remove protection if exists - WITH PASSWORD
  if (sheet.getProtection().getProtected()) {
    sheet.getProtection().unprotect(currentPassword);
  }

  // Unlock everything first
    sheet.getRange("A:AZ").getFormat().getProtection().setLocked(false);

  // Validate column letter format
  if (!/^[A-Z]+$/.test(columnLetter)) {
    return;
  }

  // Find data in specified column and lock it
  let columnToLock = sheet.getRange(`${columnLetter}:${columnLetter}`);
  let usedRangeInColumn = columnToLock.getUsedRange(true);

  if (!usedRangeInColumn) {
    return;
  }

  usedRangeInColumn.getFormat().getProtection().setLocked(false);
  let values = usedRangeInColumn.getValues();
  let valuesLength = values.length;

  for (let i = 0; i < valuesLength; i++) {
    if (values[i][0] !== "" && values[i][0] !== null) {
      usedRangeInColumn
        .getCell(i, 0)
        .getFormat()
        .getProtection()
        .setLocked(lock);
    }
  }

  // Protect the sheet WITH PASSWORD
  sheet.getProtection().protect(
    {
      allowFormatCells: true,
      allowFormatColumns: true,
      allowFormatRows: true,
    },
    currentPassword
  );
  // const endTime = Date.now();
  // const duration = endTime - startTime;
  // console.log(`Script execution completed in ${duration}ms (${(duration / 1000).toFixed(2)}s)`);
}
