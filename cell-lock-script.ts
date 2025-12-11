function main(
  workbook: ExcelScript.Workbook,
  columnLetter: string = "E",
  password: string = "password",
  lock: boolean = true
) {
  let sheet = workbook.getActiveWorksheet();

  // Remove protection if exists - WITH PASSWORD
  if (sheet.getProtection().getProtected()) {
    sheet.getProtection().unprotect(password);
  }

  // Unlock everything first
  sheet.getRange("A:ZZ").getFormat().getProtection().setLocked(false);

  // Find data in specified column and lock it
  let columnToLock = sheet.getRange(`${columnLetter}:${columnLetter}`);
  let usedRangeInColumn = columnToLock.getUsedRange();

  if (usedRangeInColumn) {
    usedRangeInColumn.getFormat().getProtection().setLocked(lock);
    console.log(
      `Locked ${usedRangeInColumn.getRowCount()} cells in Column ${columnLetter}`
    );
  }

  // Protect the sheet WITH PASSWORD
  sheet.getProtection().protect(
    {
      allowFormatCells: true,
      allowFormatColumns: true,
      allowFormatRows: true,
    },
    password
  );

  console.log(`✅ Column ${columnLetter} cells with data are now locked`);
  console.log("🔒 Sheet is password protected");
}
