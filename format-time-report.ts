/**
 * Format the course time report without changing its values.
 * Run in Excel's Office Scripts editor (Automate > New Script).
 * The script owns/replaces conditional-formatting rules in the used Hours Spent cells.
 */
function main(
  workbook: ExcelScript.Workbook,
  sheetName: string = "report-2",
  threshold: number = 3
) {
  // 1. Choose a named sheet, not whichever tab happens to be active.
  const sheet = workbook.getWorksheet(sheetName);
  if (!sheet) throw new Error(`Cannot find worksheet "${sheetName}".`);
  if (!Number.isFinite(threshold) || threshold < 0) {
    throw new Error("Enter a nonnegative numeric threshold.");
  }

  // 2. Find this report's populated rows. Formatting alone does not extend the range.
  const report = sheet.getRange("A:D").getUsedRange(true);
  if (!report || report.getRowIndex() !== 0 || report.getColumnIndex() !== 0 ||
      report.getColumnCount() !== 4 || report.getRowCount() < 2) {
    throw new Error("Expected headers in A1:D1 and at least one data row beneath them.");
  }
  const values = report.getValues();
  const headers = ["Date", "Name", "Task", "Hours Spent"];
  if (headers.some((header, index) => values[0][index] !== header)) {
    throw new Error("Headers must be Date, Name, Task, Hours Spent in A1:D1.");
  }
  const rows = values.slice(1);
  if (rows.some(row => typeof row[0] !== "number" || !Number.isFinite(row[0]) ||
      row[0] <= 0 || typeof row[3] !== "number" || !Number.isFinite(row[3]) || row[3] < 0)) {
    throw new Error("Each data row needs a numeric Excel date and nonnegative numeric hours.");
  }

  // 3. Format all populated report rows, including the new seventh row.
  const lastRow = report.getRowIndex() + report.getRowCount();
  sheet.getRange(`A2:A${lastRow}`).setNumberFormat("m/d/yyyy");
  report.getFormat().autofitColumns();
  const hours = sheet.getRange(`D2:D${lastRow}`);
  hours.clearAllConditionalFormats();
  const rule = hours
    .addConditionalFormat(ExcelScript.ConditionalFormatType.cellValue)
    .getCellValue();
  rule.setRule({
    formula1: String(threshold),
    operator: ExcelScript.ConditionalCellValueOperator.greaterThan
  });
  rule.getFormat().getFill().setColor("#CF3338");
  rule.getFormat().getFont().setColor("#FFFFFF");

  // 4. Give us a result we can check against the worksheet.
  const highlighted = rows.filter(row => Number(row[3]) > threshold).length;
  console.log(`${sheetName}: ${rows.length} rows checked; ${highlighted} above ${threshold} hours.`);
}
