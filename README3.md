function main(workbook: ExcelScript.Workbook) {
	let selectedSheet = workbook.getActiveWorksheet();
	// Insert copied cells from MK:MK on selectedSheet to MK:MK on selectedSheet.
	selectedSheet.getRange("MK:MK").insert(ExcelScript.InsertShiftDirection.right);
	selectedSheet.getRange("MK:MK").copyFrom(selectedSheet.getRange("MK:MK"));
	// Set visibility of column(s) at range MK:MK on selectedSheet to true
	selectedSheet.getRange("MK:MK").setColumnHidden(true);
}