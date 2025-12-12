// https://learn.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.range?view=office-scripts
function main(
  workbook: ExcelScript.Workbook,
  columnLetters: string = "E,Q",
  currentPassword: string = "GreenMoon",
  lock: boolean = true,
  newPassword: string = ""
) {
  // const startTime = Date.now();
  // console.log("Script execution started");
  let sheet = workbook.getActiveWorksheet();

  // Remove protection if exists - WITH PASSWORD
  if (sheet.getProtection().getProtected()) {
    sheet.getProtection().unprotect(currentPassword);
  }

  // Unlock everything first
  //   sheet.getRange("A:AZ").getFormat().getProtection().setLocked(false);

  // Split column letters by comma and process each one
  let columnsArray = columnLetters
    .split(",")
    .map((col) => col.trim().toUpperCase());

  for (let columnLetter of columnsArray) {
    // Validate column letter format
    if (!/^[A-Z]+$/.test(columnLetter)) {
      continue;
    }

    // Find data in specified column and lock it
    let columnToLock = sheet.getRange(`${columnLetter}:${columnLetter}`);
    let usedRangeInColumn = columnToLock.getUsedRange(true);

    let values = usedRangeInColumn.getValues();
    if (usedRangeInColumn) {
      usedRangeInColumn.getFormat().getProtection().setLocked(false);
    }

    for (let i = 0; i < values.length; i++) {
      if (values[i][0] !== "" && values[i][0] !== null) {
        usedRangeInColumn
          .getCell(i, 0)
          .getFormat()
          .getProtection()
          .setLocked(lock);
      }
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
  // const endTime = Date.now();
  // const duration = endTime - startTime;
  // console.log(`Script execution completed in ${duration}ms (${(duration / 1000).toFixed(2)}s)`);
}
