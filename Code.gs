const SHEET_NAME = "ClaimsCo Emails";
const SUMMARY_SHEET_NAME = "Weekly Summary";
const HEADERS = ["ID", "Received Date", "Job ID", "ClaimsCo Rep", "Category", "Notes", "Updated At"];
const CATEGORIES = ["Inquiry", "Follow Up", "Request", "Escalation / Customer Concerns", "Others"];

function doGet(event) {
  const callback = event.parameter.callback;

  try {
    const payload = event.parameter.payload
      ? JSON.parse(event.parameter.payload)
      : { action: "list" };
    const response = handlePayload(payload);

    if (callback) {
      return jsonpResponse(callback, response);
    }

    return jsonResponse(response);
  } catch (error) {
    const response = { ok: false, error: error.message };

    if (callback) {
      return jsonpResponse(callback, response);
    }

    return jsonResponse(response);
  }
}

function doPost(event) {
  try {
    const payload = JSON.parse(event.postData.contents || "{}");
    return jsonResponse(handlePayload(payload));
  } catch (error) {
    return jsonResponse({ ok: false, error: error.message });
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("ClaimsCo")
    .addItem("Create Weekly Summary PDF", "createWeeklySummaryPdfFromPrompt")
    .addToUi();
}

function createWeeklySummaryPdfFromPrompt() {
  const ui = SpreadsheetApp.getUi();
  const fromPrompt = ui.prompt("ClaimsCo Weekly Summary", "From date (dd / mm / yyyy)", ui.ButtonSet.OK_CANCEL);

  if (fromPrompt.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const toPrompt = ui.prompt("ClaimsCo Weekly Summary", "To date (dd / mm / yyyy)", ui.ButtonSet.OK_CANCEL);

  if (toPrompt.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  try {
    const report = createPdfSummary(fromPrompt.getResponseText(), toPrompt.getResponseText());
    ui.alert(`PDF summary created:\n${report.pdfUrl}`);
  } catch (error) {
    ui.alert(`Could not create PDF summary:\n${error.message}`);
  }
}

function handlePayload(payload) {
  const action = payload.action || "list";

  if (action === "list") {
    return withSpreadsheetUrl({ ok: true, records: readRecords() });
  }

  if (action === "upsert") {
    const record = normalizeRecord(payload.record || {});
    upsertRecord(record);
    return withSpreadsheetUrl({ ok: true, record, records: readRecords() });
  }

  if (action === "delete") {
    deleteRecord(payload.id);
    return withSpreadsheetUrl({ ok: true, records: readRecords() });
  }

  if (action === "clear") {
    clearRecords();
    return withSpreadsheetUrl({ ok: true, records: [] });
  }

  if (action === "pdfSummary") {
    return withSpreadsheetUrl(Object.assign({ ok: true }, createPdfSummary(payload.from, payload.to)));
  }

  return { ok: false, error: "Unknown action." };
}

function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSpreadsheetUrl() {
  return getSpreadsheet().getUrl();
}

function withSpreadsheetUrl(payload) {
  payload.spreadsheetUrl = getSpreadsheetUrl();
  return payload;
}

function getSheet() {
  const spreadsheet = getSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  ensureHeaders(sheet);
  formatDateColumn(sheet);
  return sheet;
}

function ensureHeaders(sheet) {
  const currentHeaders = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const hasHeaders = HEADERS.every((header, index) => currentHeaders[index] === header);

  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }
}

function formatDateColumn(sheet) {
  const dataRowCount = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2, 2, dataRowCount, 1).setNumberFormat("dd / mm / yyyy");
}

function rowHasRecordData(row) {
  return row
    .slice(1, HEADERS.length)
    .some((value) => String(value || "").trim().length > 0);
}

function fillMissingRecordIds(sheet, values) {
  values.slice(1).forEach((row, index) => {
    if (!rowHasRecordData(row) || row[0]) {
      return;
    }

    const rowNumber = index + 2;
    row[0] = Utilities.getUuid();
    sheet.getRange(rowNumber, 1).setValue(row[0]);
  });

  SpreadsheetApp.flush();
}

function readRecords() {
  const sheet = getSheet();
  const values = sheet.getDataRange().getValues();
  fillMissingRecordIds(sheet, values);

  return values
    .slice(1)
    .filter(rowHasRecordData)
    .map((row) => ({
      id: String(row[0] || ""),
      receivedDate: normalizeDateForApp(row[1]),
      jobId: String(row[2] || ""),
      rep: String(row[3] || ""),
      category: CATEGORIES.includes(row[4]) ? row[4] : "Others",
      notes: String(row[5] || ""),
      updatedAt: String(row[6] || ""),
    }));
}

function pad2(value) {
  return String(value).padStart(2, "0");
}

function normalizeDateForApp(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }

  const text = String(value || "").trim();
  if (!text) {
    return "";
  }

  const isoMatch = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (isoMatch) {
    return `${isoMatch[1]}-${isoMatch[2]}-${isoMatch[3]}`;
  }

  const slashMatch = text.match(/^(\d{1,2})\s*\/\s*(\d{1,2})\s*\/\s*(\d{4})$/);
  if (slashMatch) {
    const day = Number(slashMatch[1]);
    const month = Number(slashMatch[2]);
    const year = Number(slashMatch[3]);

    if (year && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      return `${year}-${pad2(month)}-${pad2(day)}`;
    }
  }

  const parsedDate = new Date(text);
  if (!isNaN(parsedDate.getTime())) {
    return Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }

  return text;
}

function isoToSheetDate(isoDate) {
  const parts = String(isoDate || "").split("-").map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2]);
}

function getOrderedDateRange(fromValue, toValue) {
  const from = normalizeDateForApp(fromValue);
  const to = normalizeDateForApp(toValue);

  if (!/^\d{4}-\d{2}-\d{2}$/.test(from) || from < "2026-01-01") {
    throw new Error("From date must be 2026 or later.");
  }

  if (!/^\d{4}-\d{2}-\d{2}$/.test(to) || to < "2026-01-01") {
    throw new Error("To date must be 2026 or later.");
  }

  return from <= to ? { from, to } : { from: to, to: from };
}

function formatSlashDate(isoDate) {
  const parts = String(isoDate || "").split("-").map(Number);
  return `${pad2(parts[2])} / ${pad2(parts[1])} / ${parts[0]}`;
}

function formatDateForFileName(isoDate) {
  const parts = String(isoDate || "").split("-").map(Number);
  return `${pad2(parts[2])}-${pad2(parts[1])}-${parts[0]}`;
}

function getSummarySheet() {
  const spreadsheet = getSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SUMMARY_SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SUMMARY_SHEET_NAME);
  }

  return sheet;
}

function buildCategorySummary(records) {
  const total = records.length;
  return CATEGORIES.map((category) => {
    const count = records.filter((record) => record.category === category).length;
    return {
      category,
      count,
      percent: total > 0 ? count / total : 0,
    };
  });
}

function createPdfSummary(fromValue, toValue) {
  const range = getOrderedDateRange(fromValue, toValue);
  const records = readRecords()
    .filter((record) => record.receivedDate >= range.from && record.receivedDate <= range.to)
    .sort((a, b) => a.receivedDate.localeCompare(b.receivedDate));
  const summary = buildCategorySummary(records);
  const sheet = writeWeeklySummarySheet(range, records, summary);
  const pdfFile = createWeeklySummaryPdfFile(sheet, range);

  return {
    pdfUrl: pdfFile.getUrl(),
    pdfName: pdfFile.getName(),
    total: records.length,
    summary,
  };
}

function writeWeeklySummarySheet(range, records, summary) {
  const sheet = getSummarySheet();
  sheet.clearCharts();
  sheet.clear();

  sheet.getRange("A1").setValue("ClaimsCo Weekly Summary");
  sheet.getRange("A2").setValue(`${formatSlashDate(range.from)} to ${formatSlashDate(range.to)}`);
  sheet.getRange("A3:B3").setValues([["Total Emails", records.length]]);
  sheet.getRange("A5:C5").setValues([["Category", "Emails", "Share"]]);

  const summaryRows = summary.map((item) => [item.category, item.count, item.percent]);
  sheet.getRange(6, 1, summaryRows.length, 3).setValues(summaryRows);
  sheet.getRange(6, 3, summaryRows.length, 1).setNumberFormat("0%");

  const detailStartRow = 13;
  sheet.getRange(detailStartRow, 1, 1, 6).setValues([
    ["Received Date", "Job ID", "ClaimsCo Rep", "Category", "Notes", "Updated At"],
  ]);

  if (records.length > 0) {
    const detailRows = records.map((record) => [
      isoToSheetDate(record.receivedDate),
      record.jobId,
      record.rep,
      record.category,
      record.notes,
      record.updatedAt,
    ]);
    sheet.getRange(detailStartRow + 1, 1, detailRows.length, 6).setValues(detailRows);
    sheet.getRange(detailStartRow + 1, 1, detailRows.length, 1).setNumberFormat("dd / mm / yyyy");
  }

  sheet.getRange("A1:F1").setFontSize(18).setFontWeight("bold");
  sheet.getRange("A2:F3").setFontWeight("bold");
  sheet.getRange("A5:C5").setFontWeight("bold").setBackground("#eee9dd");
  sheet.getRange(detailStartRow, 1, 1, 6).setFontWeight("bold").setBackground("#eee9dd");
  sheet.getRange(1, 1, Math.max(detailStartRow + records.length, 14), 6).setVerticalAlignment("top");
  sheet.getRange(detailStartRow + 1, 5, Math.max(records.length, 1), 1).setWrap(true);
  sheet.setFrozenRows(0);
  sheet.setColumnWidths(1, 1, 120);
  sheet.setColumnWidths(2, 2, 150);
  sheet.setColumnWidth(4, 180);
  sheet.setColumnWidth(5, 330);
  sheet.setColumnWidth(6, 210);

  if (records.length > 0) {
    const chart = sheet
      .newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(sheet.getRange(5, 1, summaryRows.length + 1, 2))
      .setOption("title", "Category Summary")
      .setOption("pieHole", 0.45)
      .setPosition(5, 5, 0, 0)
      .build();
    sheet.insertChart(chart);
  }

  SpreadsheetApp.flush();
  return sheet;
}

function createWeeklySummaryPdfFile(sheet, range) {
  const spreadsheet = getSpreadsheet();
  const fileName = `ClaimsCo Weekly Summary ${formatDateForFileName(range.from)} to ${formatDateForFileName(range.to)}.pdf`;
  const exportUrl = `https://docs.google.com/spreadsheets/d/${spreadsheet.getId()}/export?format=pdf&gid=${sheet.getSheetId()}&size=A4&portrait=false&fitw=true&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false&top_margin=0.40&bottom_margin=0.40&left_margin=0.40&right_margin=0.40`;
  const response = UrlFetchApp.fetch(exportUrl, {
    headers: {
      Authorization: `Bearer ${ScriptApp.getOAuthToken()}`,
    },
  });
  const file = DriveApp.createFile(response.getBlob().setName(fileName));

  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (error) {
    // Some company Google Workspace policies do not allow public sharing.
  }

  return file;
}

function normalizeRecord(record) {
  const normalized = {
    id: String(record.id || Utilities.getUuid()),
    receivedDate: normalizeDateForApp(record.receivedDate),
    jobId: String(record.jobId || "").trim(),
    rep: String(record.rep || "").trim(),
    category: CATEGORIES.includes(record.category) ? record.category : "Others",
    notes: String(record.notes || "").trim(),
    updatedAt: String(record.updatedAt || new Date().toISOString()),
  };

  if (!/^\d{4}-\d{2}-\d{2}$/.test(normalized.receivedDate) || normalized.receivedDate < "2026-01-01") {
    throw new Error("Received Date must be 2026 or later.");
  }

  if (!normalized.jobId) {
    throw new Error("Job ID is required.");
  }

  if (!normalized.rep) {
    throw new Error("ClaimsCo Rep is required.");
  }

  return normalized;
}

function writeRecordRow(sheet, rowNumber, rowValues) {
  sheet.getRange(rowNumber, 1, 1, HEADERS.length).setValues([rowValues]);
  sheet.getRange(rowNumber, 2).setNumberFormat("dd / mm / yyyy");
}

function upsertRecord(record) {
  const sheet = getSheet();
  const values = sheet.getDataRange().getValues();
  const rowIndex = values.findIndex((row, index) => index > 0 && row[0] === record.id);
  const rowValues = [
    record.id,
    isoToSheetDate(record.receivedDate),
    record.jobId,
    record.rep,
    record.category,
    record.notes,
    record.updatedAt,
  ];

  if (rowIndex >= 0) {
    writeRecordRow(sheet, rowIndex + 1, rowValues);
    return;
  }

  sheet.appendRow(rowValues);
  sheet.getRange(sheet.getLastRow(), 2).setNumberFormat("dd / mm / yyyy");
}

function deleteRecord(id) {
  const sheet = getSheet();
  const values = sheet.getDataRange().getValues();
  const rowIndex = values.findIndex((row, index) => index > 0 && row[0] === id);

  if (rowIndex >= 0) {
    sheet.deleteRow(rowIndex + 1);
  }
}

function clearRecords() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonpResponse(callback, payload) {
  const safeCallback = String(callback || "").replace(/[^\w.$]/g, "");

  if (!safeCallback) {
    return jsonResponse({ ok: false, error: "Missing callback." });
  }

  return ContentService
    .createTextOutput(`${safeCallback}(${JSON.stringify(payload)});`)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
