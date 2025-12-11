function main(
  workbook: ExcelScript.Workbook,
  columnLetters: string = "E",
  currentPassword: string = "GreenMoon",
  lock: boolean = true,
  newPassword: string = ""
) {
  let sheet = workbook.getActiveWorksheet();

  // Remove protection if exists - WITH PASSWORD
  if (sheet.getProtection().getProtected()) {
    sheet.getProtection().unprotect(currentPassword);
  }

  // Unlock everything first
  sheet.getRange("A:ZZ").getFormat().getProtection().setLocked(false);

  // Split column letters by comma and process each one
  let columnsArray = columnLetters
    .split(",")
    .map((col) => col.trim().toUpperCase());

  for (let columnLetter of columnsArray) {
    // Validate column letter format
    if (!/^[A-Z]+$/.test(columnLetter)) {
      console.log(`⚠️ Skipping invalid column letter: ${columnLetter}`);
      continue;
    }

    // Find data in specified column and lock it
    let columnToLock = sheet.getRange(`${columnLetter}:${columnLetter}`);
    let usedRangeInColumn = columnToLock.getUsedRange();

    if (usedRangeInColumn) {
      usedRangeInColumn.getFormat().getProtection().setLocked(lock);
      console.log(
        `Locked ${usedRangeInColumn.getRowCount()} cells in Column ${columnLetter}`
      );
    } else {
      console.log(`No data found in Column ${columnLetter}`);
    }
  }

  newPassword = newPassword ? newPassword : currentPassword;

  // Protect the sheet WITH PASSWORD
  sheet.getProtection().protect(
    {
      allowFormatCells: true,
      allowFormatColumns: true,
      allowFormatRows: true,
    },
    newPassword
  );

  console.log(
    `✅ Columns ${columnLetters} are now ${lock ? "locked" : "unlocked"}`
  );
  console.log("🔒 Sheet is password protected");
}
