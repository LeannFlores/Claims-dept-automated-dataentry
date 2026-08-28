function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu("Claims Automated").addItem("Open Data Entry Form", "showDataEntryForm").addToUi();
}

function showDataEntryForm() {
  var html = HtmlService.createHtmlOutputFromFile("Form").setWidth(1100).setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, "Claims Registry - 2026 | Claims Department");
}

// MULTI-COLUMN SEARCH (Column B, D, or F)
function searchRecord(query) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("test_Monitoring_2");
  if (!sheet) return null;

  var data = sheet.getDataRange().getValues();
  var searchStr = query.toString().trim().toLowerCase();

  for (var i = 1; i < data.length; i++) {
    // Column Maps: B=Col 2 (Index 1), D=Col 4 (Index 3), F=Col 6 (Index 5)
    var colB = data[i][1] ? data[i][1].toString().trim().toLowerCase() : "";
    var colD = data[i][3] ? data[i][3].toString().trim().toLowerCase() : "";
    var colF = data[i][5] ? data[i][5].toString().trim().toLowerCase() : "";

    if (colB === searchStr || colD === searchStr || colF === searchStr) {
      return {
        rowNumber: i + 1,
        inlis: data[i][0],
        branch: data[i][1],
        dateFiled: data[i][2] instanceof Date ? data[i][2].toISOString().split("T")[0] : data[i][2],
        assured: data[i][3],
        policy: data[i][5], // Column F
        agent: data[i][9], // Column J
        dateReported: data[i][10] instanceof Date ? data[i][10].toISOString().split("T")[0] : data[i][10],
        dateLoss: data[i][11] instanceof Date ? data[i][11].toISOString().split("T")[0] : data[i][11],
        from: data[i][12] instanceof Date ? data[i][12].toISOString().split("T")[0] : data[i][12],
        to: data[i][13] instanceof Date ? data[i][13].toISOString().split("T")[0] : data[i][13],
        sumInsured: data[i][14],
        natureClaim: data[i][15],
        typeVehicle: data[i][16],
        plateNo: data[i][17],
        insuredUnit: data[i][18],
        engineNo: data[i][19],
        chassis: data[i][20],
        payee: data[i][21],
        preferredShop: data[i][22],
        lossReserved: data[i][23],
        amountY: data[i][24],
        status: data[i][26], // Column AA
        remarksDate: data[i][27], // Column AB
        dateAc: data[i][28], // Column AC
        payeeOd: data[i][29], // Column AD
        amountAf: data[i][31], // Column AF
        dateClosed: data[i][32], // Column AG
        amountPaid: data[i][33], // Column AH
        voucherNo: data[i][34], // Column AI
        orNo: data[i][35], // Column AJ
        datePaidVoucher: data[i][36], // Column AK
        paymentFrom: data[i][37], // Column AL
        dateForwarded: data[i][38] instanceof Date ? data[i][38].toISOString().split("T")[0] : data[i][38], // Column AM
        adjuster: data[i][55], // Column BD
      };
    }
  }
  return null;
}

// SAVE A NEW RECORD IN SPECIFIC COLUMNS
function processForm(d) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("test_Monitoring_2");
  if (!sheet) throw new Error("Sheet 'test_Monitoring_2' not found!");

  var lastRow = sheet.getLastRow() + 1;
  writeRowData(sheet, lastRow, d);
}

// UPDATE EXISTING RECORD IN SPECIFIC COLUMNS
function updateForm(d) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("test_Monitoring_2");
  var row = parseInt(d.rowNumber);
  if (!sheet || !row) return;

  writeRowData(sheet, row, d);
}

// Shared helper function to map exact column coordinates
// Shared helper function to map exact column coordinates with safety checks
function writeRowData(sheet, row, d) {
  if (!sheet) {
    throw new Error("Target sheet could not be initialized.");
  }

  try {
    sheet.getRange(row, 1).setValue(d.inlis || ""); // A
    sheet.getRange(row, 2).setValue(d.branch || ""); // B
    sheet.getRange(row, 3).setValue(d.dateFiled || ""); // C
    sheet.getRange(row, 4).setValue(d.assured || ""); // D
    sheet.getRange(row, 6).setValue(d.policy || ""); // F
    sheet.getRange(row, 10).setValue(d.agent || ""); // J
    sheet.getRange(row, 11).setValue(d.dateReported || ""); // K
    sheet.getRange(row, 12).setValue(d.dateLoss || ""); // L
    sheet.getRange(row, 13).setValue(d.from || ""); // M
    sheet.getRange(row, 14).setValue(d.to || ""); // N
    sheet.getRange(row, 15).setValue(d.sumInsured || ""); // O
    sheet.getRange(row, 16).setValue(d.natureClaim || ""); // P
    sheet.getRange(row, 17).setValue(d.typeVehicle || ""); // Q
    sheet.getRange(row, 18).setValue(d.plateNo || ""); // R
    sheet.getRange(row, 19).setValue(d.insuredUnit || ""); // S
    sheet.getRange(row, 20).setValue(d.engineNo || ""); // T
    sheet.getRange(row, 21).setValue(d.chassis || ""); // U
    sheet.getRange(row, 22).setValue(d.payee || ""); // V
    sheet.getRange(row, 23).setValue(d.preferredShop || ""); // W
    sheet.getRange(row, 24).setValue(d.lossReserved || ""); // X
    sheet.getRange(row, 25).setValue(d.amountY || ""); // Y
    sheet.getRange(row, 27).setValue(d.status || ""); // AA
    sheet.getRange(row, 28).setValue(d.remarksDate || ""); // AB
    sheet.getRange(row, 29).setValue(d.dateAc || ""); // AC
    sheet.getRange(row, 30).setValue(d.payeeOd || ""); // AD
    sheet.getRange(row, 32).setValue(d.amountAf || ""); // AF
    sheet.getRange(row, 33).setValue(d.dateClosed || ""); // AG
    sheet.getRange(row, 34).setValue(d.amountPaid || ""); // AH
    sheet.getRange(row, 35).setValue(d.voucherNo || ""); // AI
    sheet.getRange(row, 36).setValue(d.orNo || ""); // AJ
    sheet.getRange(row, 37).setValue(d.datePaidVoucher || ""); // AK
    sheet.getRange(row, 38).setValue(d.paymentFrom || ""); // AL
    sheet.getRange(row, 39).setValue(d.dateForwarded || ""); // AM
    sheet.getRange(row, 56).setValue(d.adjuster || ""); // BD
  } catch (err) {
    throw new Error("Failed to write cell data: " + err.message);
  }
}
