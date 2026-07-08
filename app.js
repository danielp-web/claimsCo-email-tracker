const minimumYear = 2026;
const storageKey = "claimsco-email-tracker-records";
const categoryColumns = [
  "Inquiry",
  "Follow Up",
  "Request",
  "Escalation / Customer Concerns",
  "Others",
];
const categoryColors = ["#146c66", "#c98721", "#5c77a8", "#b94a3c", "#7c5fa6"];
const monthLabels = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const appConfig = window.CLAIMSCO_CONFIG || {};
const sheetsWebAppUrl = String(appConfig.SHEETS_WEB_APP_URL || "").trim();
const configuredSpreadsheetUrl = String(appConfig.SPREADSHEET_URL || "").trim();
const isSharedMode = sheetsWebAppUrl.length > 0;
const allowSharedRowChanges = appConfig.ALLOW_SHARED_ROW_CHANGES === true;
const autoRefreshSeconds = Math.max(5, Number(appConfig.AUTO_REFRESH_SECONDS) || 10);

const form = document.querySelector("#entryForm");
const submitButton = form.querySelector("button[type='submit']");
const tableBody = document.querySelector("#tableBody");
const emptyState = document.querySelector("#emptyState");
const clearButton = document.querySelector("#clearButton");
const resetButton = document.querySelector("#resetButton");
const pdfSummaryButton = document.querySelector("#pdfSummaryButton");
const pdfStatus = document.querySelector("#pdfStatus");
const rowCount = document.querySelector("#rowCount");
const emailTotal = document.querySelector("#emailTotal");
const syncStatus = document.querySelector("#syncStatus");
const spreadsheetLink = document.querySelector("#spreadsheetLink");
const viewButtons = Array.from(document.querySelectorAll("[data-view-mode]"));
const viewStatus = document.querySelector("#viewStatus");
const submitLabel = document.querySelector("#submitLabel");
const editingId = document.querySelector("#editingId");
const emptyTitle = document.querySelector("#emptyTitle");
const emptyMessage = document.querySelector("#emptyMessage");

const fields = {
  receivedDay: document.querySelector("#receivedDayInput"),
  receivedMonth: document.querySelector("#receivedMonthInput"),
  receivedYear: document.querySelector("#receivedYearInput"),
  jobId: document.querySelector("#jobIdInput"),
  rep: document.querySelector("#repInput"),
  category: document.querySelector("#categoryInput"),
  notes: document.querySelector("#notesInput"),
};

const dateControls = {
  received: {
    day: fields.receivedDay,
    month: fields.receivedMonth,
    year: fields.receivedYear,
    affectsReport: false,
  },
  summaryFrom: {
    day: document.querySelector("#summaryFromDayInput"),
    month: document.querySelector("#summaryFromMonthInput"),
    year: document.querySelector("#summaryFromYearInput"),
    affectsReport: false,
  },
  summaryTo: {
    day: document.querySelector("#summaryToDayInput"),
    month: document.querySelector("#summaryToMonthInput"),
    year: document.querySelector("#summaryToYearInput"),
    affectsReport: false,
  },
};

let records = loadRecords();
let activeViewMode = "all";
let isRefreshingSharedRecords = false;
let sharedAutoRefreshTimer = null;

function pad2(value) {
  return String(value).padStart(2, "0");
}

function daysInMonth(year, month) {
  return new Date(year, month, 0).getDate();
}

function localDateToIso(date) {
  return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
}

function parseDateToIso(value) {
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
    const safeDay = Math.min(day, daysInMonth(year, month));

    if (year && month >= 1 && month <= 12 && safeDay >= 1) {
      return `${year}-${pad2(month)}-${pad2(safeDay)}`;
    }
  }

  const parsedDate = new Date(text);
  if (!Number.isNaN(parsedDate.getTime())) {
    return localDateToIso(parsedDate);
  }

  return text;
}

function isoToLocalDate(value) {
  const isoDate = parseDateToIso(value);
  const [year, month, day] = String(isoDate || "").split("-").map(Number);
  if (!year || !month || !day) {
    return null;
  }
  return new Date(year, month - 1, day);
}

function todayIsoFrom2026() {
  const today = new Date();
  if (today.getFullYear() < minimumYear) {
    return `${minimumYear}-01-01`;
  }
  return localDateToIso(today);
}

function clampTo2026(isoDate) {
  const safeDate = parseDateToIso(isoDate);
  if (!safeDate || safeDate < `${minimumYear}-01-01`) {
    return `${minimumYear}-01-01`;
  }
  return safeDate;
}

function formatSlashDate(isoDate) {
  const date = isoToLocalDate(isoDate);
  if (!date) {
    return "No date";
  }

  return `${pad2(date.getDate())} / ${pad2(date.getMonth() + 1)} / ${date.getFullYear()}`;
}

function formatDisplayDate(isoDate) {
  return formatSlashDate(isoDate);
}

function isWeekend(isoDate) {
  const date = isoToLocalDate(isoDate);
  if (!date) {
    return false;
  }
  const day = date.getDay();
  return day === 0 || day === 6;
}

function populateMonthOptions(select) {
  select.innerHTML = monthLabels
    .map((label, index) => `<option value="${index + 1}">${label}</option>`)
    .join("");
}

function syncDayOptions(controlName) {
  const control = dateControls[controlName];
  const year = Math.max(Number(control.year.value) || minimumYear, minimumYear);
  const month = Number(control.month.value) || 1;
  const selectedDay = Math.min(Number(control.day.value) || 1, daysInMonth(year, month));
  const dayCount = daysInMonth(year, month);

  control.day.innerHTML = Array.from({ length: dayCount }, (_, index) => {
    const day = index + 1;
    return `<option value="${day}">${day}</option>`;
  }).join("");
  control.day.value = String(selectedDay);
}

function getDateControlIso(controlName) {
  const control = dateControls[controlName];
  const year = Math.max(Number(control.year.value) || minimumYear, minimumYear);
  const month = Number(control.month.value) || 1;
  const day = Number(control.day.value) || 1;
  const safeDay = Math.min(day, daysInMonth(year, month));

  return `${year}-${pad2(month)}-${pad2(safeDay)}`;
}

function setDateControl(controlName, isoDate) {
  const control = dateControls[controlName];
  const safeDate = clampTo2026(isoDate);
  const [year, month, day] = safeDate.split("-").map(Number);

  control.year.value = year;
  control.month.value = month;
  syncDayOptions(controlName);
  control.day.value = String(day);
}

function setupDateControl(controlName) {
  const control = dateControls[controlName];
  populateMonthOptions(control.month);
  control.year.min = minimumYear;

  const handleChange = () => {
    syncDayOptions(controlName);
    if (control.affectsReport) {
      render();
    }
  };

  control.day.addEventListener("change", () => {
    if (control.affectsReport) {
      render();
    }
  });
  control.month.addEventListener("change", handleChange);
  control.year.addEventListener("input", handleChange);
}

function startOfWorkWeek(isoDate = todayIsoFrom2026()) {
  const date = isoToLocalDate(isoDate);
  const day = date.getDay();
  const mondayOffset = day === 0 ? -6 : 1 - day;
  date.setDate(date.getDate() + mondayOffset);
  return clampTo2026(localDateToIso(date));
}

function endOfWorkWeek(isoDate = todayIsoFrom2026()) {
  const start = isoToLocalDate(startOfWorkWeek(isoDate));
  start.setDate(start.getDate() + 4);
  return localDateToIso(start);
}

function startOfMonth(isoDate = todayIsoFrom2026()) {
  const date = isoToLocalDate(isoDate);
  return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-01`;
}

function endOfMonth(isoDate = todayIsoFrom2026()) {
  const date = isoToLocalDate(isoDate);
  return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(daysInMonth(date.getFullYear(), date.getMonth() + 1))}`;
}

function normalizeRecord(record) {
  return {
    id: record.id || crypto.randomUUID(),
    receivedDate: parseDateToIso(record.receivedDate || ""),
    jobId: record.jobId || record.name || "",
    rep: record.rep || record.email || "",
    category: categoryColumns.includes(record.category) ? record.category : "Others",
    notes: record.notes || "",
    updatedAt: record.updatedAt || new Date().toISOString(),
  };
}

function loadRecords() {
  try {
    const saved = localStorage.getItem(storageKey);
    return saved ? JSON.parse(saved).map(normalizeRecord) : [];
  } catch {
    return [];
  }
}

function saveRecords() {
  localStorage.setItem(storageKey, JSON.stringify(records));
}

function canChangeExistingRows() {
  return !isSharedMode || allowSharedRowChanges;
}

function setSyncStatus(message, tone = "local") {
  if (!syncStatus) {
    return;
  }

  syncStatus.textContent = message;
  syncStatus.dataset.tone = tone;
}

function setSpreadsheetLink(url) {
  if (!spreadsheetLink) {
    return;
  }

  const safeUrl = String(url || "").trim();
  spreadsheetLink.classList.toggle("is-hidden", !safeUrl);

  if (safeUrl) {
    spreadsheetLink.href = safeUrl;
  }
}

function setSubmitBusy(isBusy) {
  submitButton.disabled = isBusy;
  submitButton.classList.toggle("is-busy", isBusy);
}

function setPdfStatus(message, tone = "local") {
  if (!pdfStatus) {
    return;
  }

  pdfStatus.textContent = message;
  pdfStatus.dataset.tone = tone;
}

function setPdfBusy(isBusy) {
  if (!pdfSummaryButton) {
    return;
  }

  pdfSummaryButton.disabled = isBusy;
  pdfSummaryButton.classList.toggle("is-busy", isBusy);
}

async function sharedRequest(payload) {
  return new Promise((resolve, reject) => {
    const callbackName = `claimscoCallback_${Date.now()}_${Math.random().toString(36).slice(2)}`;
    const script = document.createElement("script");
    const timeout = window.setTimeout(() => {
      cleanup();
      reject(new Error("Spreadsheet request timed out. Check Apps Script deployment access."));
    }, 20000);

    function cleanup() {
      window.clearTimeout(timeout);
      delete window[callbackName];
      script.remove();
    }

    window[callbackName] = (data) => {
      cleanup();
      if (!data || data.ok === false) {
        reject(new Error(data?.error || "Spreadsheet request failed."));
        return;
      }
      resolve(data);
    };

    const url = new URL(sheetsWebAppUrl);
    url.searchParams.set("callback", callbackName);
    url.searchParams.set("payload", JSON.stringify(payload));
    script.src = url.toString();
    script.onerror = () => {
      cleanup();
      reject(
        new Error(
          "Spreadsheet script could not load. Redeploy the latest Code.gs as a new Web app version and set access to Anyone with the link.",
        ),
      );
    };
    document.head.append(script);
  });
}

async function refreshSharedRecords(options = {}) {
  if (!isSharedMode) {
    setSyncStatus("Local browser storage", "local");
    return;
  }

  if (isRefreshingSharedRecords) {
    return;
  }

  isRefreshingSharedRecords = true;

  if (!options.quiet) {
    setSyncStatus("Loading shared spreadsheet...", "loading");
  }

  try {
    const data = await sharedRequest({ action: "list" });
    records = (data.records || []).map(normalizeRecord);
    saveRecords();
    render();
    setSpreadsheetLink(data.spreadsheetUrl || configuredSpreadsheetUrl);
    setSyncStatus("Shared spreadsheet connected - auto-updating", "connected");
  } catch (error) {
    console.error(error);
    setSyncStatus(`${error.message} Showing local copy.`, "error");
  } finally {
    isRefreshingSharedRecords = false;
  }
}

function startSharedAutoRefresh() {
  if (!isSharedMode || sharedAutoRefreshTimer) {
    return;
  }

  sharedAutoRefreshTimer = window.setInterval(() => {
    if (document.visibilityState === "visible") {
      refreshSharedRecords({ quiet: true });
    }
  }, autoRefreshSeconds * 1000);

  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") {
      refreshSharedRecords({ quiet: true });
    }
  });

  window.addEventListener("focus", () => {
    refreshSharedRecords({ quiet: true });
  });
}

function configureSharedModeUi() {
  if (!isSharedMode) {
    clearButton.classList.remove("is-hidden");
    return;
  }

  clearButton.classList.toggle("is-hidden", !allowSharedRowChanges);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function getVisibleRecords() {
  const range = getActiveViewRange();

  if (!range) {
    return records;
  }

  return getRecordsForRange(range.from, range.to);
}

function getSummaryDateRange() {
  const from = getDateControlIso("summaryFrom");
  const to = getDateControlIso("summaryTo");
  return from <= to ? { from, to } : { from: to, to: from };
}

function getActiveViewRange() {
  const today = todayIsoFrom2026();

  if (activeViewMode === "week") {
    return {
      from: startOfWorkWeek(today),
      to: endOfWorkWeek(today),
    };
  }

  if (activeViewMode === "month") {
    return {
      from: startOfMonth(today),
      to: endOfMonth(today),
    };
  }

  return null;
}

function getRecordsForRange(from, to) {
  return records.filter((record) => {
    const receivedDate = parseDateToIso(record.receivedDate);
    return receivedDate >= from && receivedDate <= to;
  });
}

function getViewStatusText(visibleCount) {
  const range = getActiveViewRange();

  if (!range) {
    return "Showing all saved emails";
  }

  const label = activeViewMode === "week" ? "current week" : "current month";
  return `Showing ${visibleCount} emails for the ${label}: ${formatDisplayDate(range.from)} to ${formatDisplayDate(range.to)}`;
}

function syncSummaryRangeToView() {
  const range = getActiveViewRange();

  if (!range) {
    return;
  }

  setDateControl("summaryFrom", range.from);
  setDateControl("summaryTo", range.to);
}

function updateViewControls(visibleCount) {
  viewButtons.forEach((button) => {
    const isActive = button.dataset.viewMode === activeViewMode;
    button.classList.toggle("is-active", isActive);
    button.setAttribute("aria-pressed", String(isActive));
  });

  if (viewStatus) {
    viewStatus.textContent = getViewStatusText(visibleCount);
  }
}

function setTableViewMode(mode) {
  activeViewMode = ["all", "week", "month"].includes(mode) ? mode : "all";
  syncSummaryRangeToView();
  render();
}

function renderCategoryCell(record, category) {
  if (record.category !== category) {
    return "";
  }

  return `<span class="category-mark" aria-label="${escapeHtml(category)}">Yes</span>`;
}

function getCategorySummary(rows) {
  const total = rows.length;
  return categoryColumns.map((category, index) => {
    const count = rows.filter((record) => record.category === category).length;
    return {
      category,
      color: categoryColors[index],
      count,
      percent: total > 0 ? Math.round((count / total) * 100) : 0,
    };
  });
}

function polarToCartesian(cx, cy, radius, angleInDegrees) {
  const angleInRadians = ((angleInDegrees - 90) * Math.PI) / 180;
  return {
    x: cx + radius * Math.cos(angleInRadians),
    y: cy + radius * Math.sin(angleInRadians),
  };
}

function describeArc(cx, cy, radius, startAngle, endAngle) {
  const start = polarToCartesian(cx, cy, radius, endAngle);
  const end = polarToCartesian(cx, cy, radius, startAngle);
  const largeArcFlag = endAngle - startAngle <= 180 ? "0" : "1";

  return [
    `M ${cx} ${cy}`,
    `L ${start.x} ${start.y}`,
    `A ${radius} ${radius} 0 ${largeArcFlag} 0 ${end.x} ${end.y}`,
    "Z",
  ].join(" ");
}

function buildPieChartSvg(summary, total) {
  if (total === 0) {
    return `
      <svg class="pie-chart" viewBox="0 0 220 220" role="img" aria-label="No emails chart">
        <circle cx="110" cy="110" r="88" fill="#eee9dd"></circle>
        <circle cx="110" cy="110" r="54" fill="#fbfaf6"></circle>
        <text x="110" y="105" text-anchor="middle" class="pie-total">0</text>
        <text x="110" y="128" text-anchor="middle" class="pie-label">emails</text>
      </svg>
    `;
  }

  let currentAngle = 0;
  const slices = summary
    .filter((item) => item.count > 0)
    .map((item) => {
      const angle = (item.count / total) * 360;
      const path = describeArc(110, 110, 88, currentAngle, currentAngle + angle);
      currentAngle += angle;
      return `<path d="${path}" fill="${item.color}"></path>`;
    })
    .join("");

  return `
    <svg class="pie-chart" viewBox="0 0 220 220" role="img" aria-label="Email category pie chart">
      ${slices}
      <circle cx="110" cy="110" r="54" fill="#fbfaf6"></circle>
      <text x="110" y="105" text-anchor="middle" class="pie-total">${total}</text>
      <text x="110" y="128" text-anchor="middle" class="pie-label">emails</text>
    </svg>
  `;
}

function buildLegendHtml(summary) {
  return summary
    .map(
      (item) => `
        <div class="legend-item">
          <span class="legend-swatch" style="background:${item.color}"></span>
          <span>${escapeHtml(item.category)}</span>
          <strong>${item.count} (${item.percent}%)</strong>
        </div>
      `,
    )
    .join("");
}

function render() {
  const visibleRecords = getVisibleRecords();
  tableBody.innerHTML = visibleRecords
    .map(
      (record) => `
        <tr>
          <td class="date-cell">${escapeHtml(formatDisplayDate(record.receivedDate))}</td>
          <td>${escapeHtml(record.jobId)}</td>
          <td>${escapeHtml(record.rep)}</td>
          ${categoryColumns
            .map((category) => `<td>${renderCategoryCell(record, category)}</td>`)
            .join("")}
          <td>${escapeHtml(record.notes)}</td>
          ${
            canChangeExistingRows()
              ? `<td>
                  <div class="row-actions">
                    <button class="icon-button" type="button" data-action="edit" data-id="${record.id}" title="Edit row" aria-label="Edit ${escapeHtml(record.jobId)}">
                      <svg aria-hidden="true" viewBox="0 0 24 24">
                        <path d="M12 20h9M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4Z" />
                      </svg>
                    </button>
                    <button class="icon-button danger" type="button" data-action="delete" data-id="${record.id}" title="Delete row" aria-label="Delete ${escapeHtml(record.jobId)}">
                      <svg aria-hidden="true" viewBox="0 0 24 24">
                        <path d="M3 6h18M8 6V4h8v2m-6 5v6m4-6v6M6 6l1 15h10l1-15" />
                      </svg>
                    </button>
                  </div>
                </td>`
              : `<td class="shared-readonly">Shared</td>`
          }
        </tr>
      `,
    )
    .join("");

  const countText = records.length === 1 ? "1 saved email" : `${records.length} saved emails`;
  rowCount.textContent = countText;
  emailTotal.textContent = visibleRecords.length;
  updateViewControls(visibleRecords.length);

  const shouldShowEmpty = visibleRecords.length === 0;
  emptyTitle.textContent = records.length > 0 ? "No emails for this view" : "No emails yet";
  emptyMessage.textContent =
    records.length > 0 ? "Use All, Week, or Month to change the list." : "Add your first ClaimsCo email with the form.";
  emptyState.classList.toggle("is-visible", shouldShowEmpty);
  tableBody.closest(".table-wrap").style.display = shouldShowEmpty ? "none" : "block";
}

function resetForm() {
  form.reset();
  editingId.value = "";
  setDateControl("received", todayIsoFrom2026());
  submitLabel.textContent = "Add row";
  fields.jobId.focus();
}

function readFormRecord() {
  return {
    id: editingId.value || crypto.randomUUID(),
    receivedDate: getDateControlIso("received"),
    jobId: fields.jobId.value.trim(),
    rep: fields.rep.value.trim(),
    category: fields.category.value,
    notes: fields.notes.value.trim(),
    updatedAt: new Date().toISOString(),
  };
}

function editRecord(id) {
  const record = records.find((item) => item.id === id);
  if (!record) {
    return;
  }

  editingId.value = record.id;
  setDateControl("received", record.receivedDate || todayIsoFrom2026());
  fields.jobId.value = record.jobId;
  fields.rep.value = record.rep;
  fields.category.value = record.category;
  fields.notes.value = record.notes;
  submitLabel.textContent = "Save row";
  fields.jobId.focus();
}

async function deleteRecord(id) {
  const previousRecords = [...records];
  records = records.filter((record) => record.id !== id);
  saveRecords();
  render();

  if (!isSharedMode) {
    return;
  }

  setSyncStatus("Deleting from shared spreadsheet...", "loading");

  try {
    await sharedRequest({ action: "delete", id });
    await refreshSharedRecords({ quiet: true });
  } catch (error) {
    console.error(error);
    records = previousRecords;
    saveRecords();
    render();
    setSyncStatus("Delete failed. Shared spreadsheet was not changed.", "error");
  }
}

function buildReportTable(rows) {
  if (rows.length === 0) {
    return `<p class="empty-report">No emails matched this report range.</p>`;
  }

  return `
    <table>
      <thead>
        <tr>
          <th>Received Date</th>
          <th>Job ID</th>
          <th>ClaimsCo Rep</th>
          ${categoryColumns.map((category) => `<th>${escapeHtml(category)}</th>`).join("")}
          <th>Notes</th>
        </tr>
      </thead>
      <tbody>
        ${rows
          .map(
            (record) => `
              <tr>
                <td>${escapeHtml(formatDisplayDate(record.receivedDate))}</td>
                <td>${escapeHtml(record.jobId)}</td>
                <td>${escapeHtml(record.rep)}</td>
                ${categoryColumns
                  .map((category) => `<td>${record.category === category ? "Yes" : ""}</td>`)
                  .join("")}
                <td>${escapeHtml(record.notes)}</td>
              </tr>
            `,
          )
          .join("")}
      </tbody>
    </table>
  `;
}

function buildReportHtml(rows, range = {}) {
  const from = range.from || rows.at(-1)?.receivedDate || todayIsoFrom2026();
  const to = range.to || rows[0]?.receivedDate || todayIsoFrom2026();
  const summary = getCategorySummary(rows);
  const weekdayText = `${formatDisplayDate(from)} to ${formatDisplayDate(to)}`;

  return `<!doctype html>
    <html lang="en">
      <head>
        <meta charset="utf-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title>ClaimsCo Email Report</title>
        <style>
          body {
            margin: 0;
            background: #f6f3ed;
            color: #20201d;
            font-family: Arial, sans-serif;
          }
          main {
            width: min(1120px, calc(100% - 36px));
            margin: 0 auto;
            padding: 28px 0 44px;
          }
          header {
            display: flex;
            justify-content: space-between;
            gap: 18px;
            align-items: flex-start;
            margin-bottom: 20px;
            border-bottom: 3px solid #146c66;
            padding-bottom: 16px;
          }
          h1, h2, p {
            margin: 0;
          }
          h1 {
            font-size: 34px;
          }
          h2 {
            font-size: 18px;
            margin-bottom: 10px;
          }
          .meta {
            color: #69665d;
            line-height: 1.6;
            text-align: right;
          }
          .summary {
            display: grid;
            grid-template-columns: 330px 1fr;
            gap: 18px;
            align-items: center;
            margin: 20px 0;
            padding: 20px;
            border: 1px solid #d8d2c6;
            border-radius: 8px;
            background: #fffefa;
          }
          .chart-box {
            display: grid;
            place-items: center;
          }
          .pie-chart {
            width: 260px;
            height: 260px;
          }
          .pie-total {
            font-size: 34px;
            font-weight: 800;
            fill: #20201d;
          }
          .pie-label {
            font-size: 14px;
            font-weight: 700;
            fill: #69665d;
          }
          .legend {
            display: grid;
            gap: 10px;
          }
          .legend-item {
            display: grid;
            grid-template-columns: 18px 1fr auto;
            gap: 10px;
            align-items: center;
            padding: 10px 12px;
            border: 1px solid #d8d2c6;
            border-radius: 8px;
            background: #fbfaf6;
          }
          .legend-swatch {
            width: 14px;
            height: 14px;
            border-radius: 999px;
          }
          .total {
            font-size: 42px;
            font-weight: 800;
            color: #146c66;
          }
          table {
            width: 100%;
            border-collapse: collapse;
            background: #fffefa;
            border: 1px solid #d8d2c6;
          }
          th, td {
            border-bottom: 1px solid #d8d2c6;
            padding: 10px;
            text-align: left;
            vertical-align: top;
            font-size: 13px;
          }
          th {
            background: #eee9dd;
            text-transform: uppercase;
            font-size: 11px;
          }
          .empty-report {
            padding: 18px;
            border: 1px solid #d8d2c6;
            border-radius: 8px;
            background: #fffefa;
          }
          @media print {
            body {
              background: #ffffff;
            }
            main {
              width: 100%;
            }
          }
        </style>
      </head>
      <body>
        <main>
          <header>
            <div>
              <h1>ClaimsCo Email Tracker Report</h1>
              <p>${escapeHtml(formatDisplayDate(from))} to ${escapeHtml(formatDisplayDate(to))}</p>
            </div>
            <div class="meta">
              <p>${escapeHtml(weekdayText)}</p>
              <p>Generated ${escapeHtml(formatDisplayDate(todayIsoFrom2026()))}</p>
            </div>
          </header>

          <section class="summary">
            <div class="chart-box">
              ${buildPieChartSvg(summary, rows.length)}
            </div>
            <div>
              <h2>Received Email Summary</h2>
              <p class="total">${rows.length}</p>
              <div class="legend">
                ${buildLegendHtml(summary)}
              </div>
            </div>
          </section>

          <section>
            <h2>Report Details</h2>
            ${buildReportTable(rows)}
          </section>
        </main>
      </body>
    </html>`;
}

function exportReport() {
  const rows = getVisibleRecords();
  const from = rows.at(-1)?.receivedDate || todayIsoFrom2026();
  const to = rows[0]?.receivedDate || todayIsoFrom2026();
  const reportHtml = buildReportHtml(rows, { from, to });
  const blob = new Blob([reportHtml], { type: "text/html;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `claimsco-email-report-${from}-to-${to}.html`;
  document.body.append(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function openPrintableSummary(rows, range, existingWindow) {
  const reportHtml = buildReportHtml(rows, range);
  const reportWindow = existingWindow || window.open("", "_blank");

  if (!reportWindow) {
    setPdfStatus("Pop-up blocked. Please allow pop-ups, then try again.", "error");
    return;
  }

  reportWindow.document.open();
  reportWindow.document.write(reportHtml);
  reportWindow.document.close();
  reportWindow.focus();
  window.setTimeout(() => reportWindow.print(), 400);
}

async function createPdfSummary() {
  const range = getSummaryDateRange();
  const localRows = getRecordsForRange(range.from, range.to);
  const reportWindow = window.open("", "_blank");

  if (reportWindow) {
    reportWindow.document.write("<p style=\"font-family:Arial,sans-serif;padding:24px\">Creating PDF summary...</p>");
  }

  setPdfBusy(true);
  setPdfStatus("Creating PDF summary...", "loading");

  try {
    if (isSharedMode) {
      const data = await sharedRequest({
        action: "pdfSummary",
        from: range.from,
        to: range.to,
      });

      if (data.pdfUrl) {
        if (reportWindow) {
          reportWindow.location.href = data.pdfUrl;
        } else {
          window.open(data.pdfUrl, "_blank", "noopener");
        }

        setPdfStatus("PDF summary ready.", "connected");
        return;
      }
    }

    openPrintableSummary(localRows, range, reportWindow);
    setPdfStatus("Printable PDF summary ready.", "connected");
  } catch (error) {
    console.error(error);
    openPrintableSummary(localRows, range, reportWindow);
    setPdfStatus("Spreadsheet PDF failed. Printable summary opened instead.", "error");
  } finally {
    setPdfBusy(false);
  }
}

Object.keys(dateControls).forEach(setupDateControl);
setDateControl("received", todayIsoFrom2026());
setDateControl("summaryFrom", startOfWorkWeek(todayIsoFrom2026()));
setDateControl("summaryTo", endOfWorkWeek(todayIsoFrom2026()));
setSpreadsheetLink(configuredSpreadsheetUrl);
render();

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  const record = readFormRecord();

  if (!record.jobId) {
    fields.jobId.focus();
    return;
  }

  if (!record.rep) {
    fields.rep.focus();
    return;
  }

  const previousRecords = [...records];
  const existingIndex = records.findIndex((item) => item.id === record.id);
  if (existingIndex >= 0) {
    records[existingIndex] = record;
  } else {
    records.unshift(record);
  }

  saveRecords();
  resetForm();
  render();

  if (!isSharedMode) {
    return;
  }

  setSubmitBusy(true);
  setSyncStatus("Saving to shared spreadsheet...", "loading");

  try {
    await sharedRequest({ action: "upsert", record });
    await refreshSharedRecords({ quiet: true });
  } catch (error) {
    console.error(error);
    records = previousRecords;
    saveRecords();
    render();
    setSyncStatus("Save failed. Please check the spreadsheet connection.", "error");
  } finally {
    setSubmitBusy(false);
  }
});

tableBody.addEventListener("click", (event) => {
  const button = event.target.closest("button[data-action]");
  if (!button) {
    return;
  }

  const { action, id } = button.dataset;
  if (action === "edit") {
    editRecord(id);
  }
  if (action === "delete") {
    deleteRecord(id);
  }
});

resetButton.addEventListener("click", resetForm);
pdfSummaryButton.addEventListener("click", createPdfSummary);
viewButtons.forEach((button) => {
  button.addEventListener("click", () => {
    setTableViewMode(button.dataset.viewMode);
  });
});
clearButton.addEventListener("click", async () => {
  if (records.length === 0) {
    return;
  }

  const message = isSharedMode
    ? "Clear all rows from the shared spreadsheet for everyone?"
    : "Clear all rows from this table?";
  const confirmed = window.confirm(message);
  if (!confirmed) {
    return;
  }

  const previousRecords = [...records];
  records = [];
  saveRecords();
  resetForm();
  render();

  if (!isSharedMode) {
    return;
  }

  setSyncStatus("Clearing shared spreadsheet...", "loading");

  try {
    await sharedRequest({ action: "clear" });
    await refreshSharedRecords({ quiet: true });
  } catch (error) {
    console.error(error);
    records = previousRecords;
    saveRecords();
    render();
    setSyncStatus("Clear failed. Shared spreadsheet was not changed.", "error");
  }
});

configureSharedModeUi();
refreshSharedRecords({ quiet: true });
startSharedAutoRefresh();
