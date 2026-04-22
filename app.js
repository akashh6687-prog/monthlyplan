const STORAGE_KEY = "iti-monthly-plan-maker-v2";
const EXPORT_SERVICE_ORIGIN = "http://127.0.0.1:3210";
const DEFAULT_STATE = {
  tradeValue: "D/CIVIL  2025-27    ( JUNIORS )  (II nd Shift)",
  monthValue: "2026-02",
  skillData: "",
  knowledgeData: "",
  overrides: [],
};

const elements = {
  tradeInput: document.getElementById("tradeInput"),
  monthInput: document.getElementById("monthInput"),
  overrideDate: document.getElementById("overrideDate"),
  overrideType: document.getElementById("overrideType"),
  overrideLabel: document.getElementById("overrideLabel"),
  addOverrideButton: document.getElementById("addOverrideButton"),
  overrideList: document.getElementById("overrideList"),
  skillDataInput: document.getElementById("skillDataInput"),
  knowledgeDataInput: document.getElementById("knowledgeDataInput"),
  renderButton: document.getElementById("renderButton"),
  printButton: document.getElementById("printButton"),
  resetButton: document.getElementById("resetButton"),
  downloadExcelButton: document.getElementById("downloadExcelButton"),
  statsPanel: document.getElementById("statsPanel"),
  tradeHeadingText: document.getElementById("tradeHeadingText"),
  monthHeadingText: document.getElementById("monthHeadingText"),
  planBody: document.getElementById("planBody"),
  planPage: document.getElementById("planPage"),
  exportStatus: document.getElementById("exportStatus"),
};

let state = loadState();
let currentPlan = null;
let exportServiceReady = false;

bootstrap();

function bootstrap() {
  syncFormWithState();
  bindEvents();
  syncOverrideDateOptions();
  renderEverything();
  checkExportService();
}

function bindEvents() {
  [
    elements.tradeInput,
    elements.monthInput,
    elements.skillDataInput,
    elements.knowledgeDataInput,
  ].forEach((field) => {
    field.addEventListener("input", handleFormChange);
  });

  elements.monthInput.addEventListener("change", () => {
    handleFormChange();
    syncOverrideDateOptions();
    renderOverrideList();
  });

  elements.renderButton.addEventListener("click", renderEverything);
  elements.printButton.addEventListener("click", () => {
    renderEverything();
    setExportStatus("info", "Browser print opened with print-color styling enabled.");
    window.print();
  });
  elements.downloadExcelButton.addEventListener("click", downloadExcel);

  elements.resetButton.addEventListener("click", () => {
    const confirmed = window.confirm("Reset all inputs and remove saved data for this planner?");
    if (!confirmed) {
      return;
    }

    state = structuredClone(DEFAULT_STATE);
    persistState();
    syncFormWithState();
    syncOverrideDateOptions();
    clearExportStatus();
    renderEverything();
  });

  elements.addOverrideButton.addEventListener("click", addOrUpdateOverride);

  elements.overrideList.addEventListener("click", (event) => {
    const button = event.target.closest("button[data-date]");
    if (!button) {
      return;
    }

    state.overrides = state.overrides.filter((override) => override.date !== button.dataset.date);
    persistState();
    renderEverything();
  });

  window.addEventListener("beforeprint", renderEverything);
  window.addEventListener("focus", checkExportService);
}

function handleFormChange() {
  state.tradeValue = elements.tradeInput.value;
  state.monthValue = elements.monthInput.value || DEFAULT_STATE.monthValue;
  state.skillData = elements.skillDataInput.value;
  state.knowledgeData = elements.knowledgeDataInput.value;

  persistState();
  renderEverything();
}

function syncFormWithState() {
  elements.tradeInput.value = state.tradeValue;
  elements.monthInput.value = state.monthValue;
  elements.skillDataInput.value = state.skillData;
  elements.knowledgeDataInput.value = state.knowledgeData;
}

function syncOverrideDateOptions() {
  const dates = getMonthDates(state.monthValue);
  const selectedValue = elements.overrideDate.value;
  elements.overrideDate.replaceChildren();

  dates.forEach((date) => {
    const option = document.createElement("option");
    option.value = formatIsoDate(date);
    option.textContent = formatDisplayDate(date);
    elements.overrideDate.append(option);
  });

  if (!dates.length) {
    return;
  }

  const fallback = formatIsoDate(dates[0]);
  elements.overrideDate.value = dates.some((date) => formatIsoDate(date) === selectedValue)
    ? selectedValue
    : fallback;
}

function addOrUpdateOverride() {
  const date = elements.overrideDate.value;
  const type = elements.overrideType.value;
  const label = elements.overrideLabel.value.trim();

  if (!date || !label) {
    window.alert("Choose a date and enter the row text first.");
    return;
  }

  const nextOverride = { date, type, label };
  const existingIndex = state.overrides.findIndex((override) => override.date === date);

  if (existingIndex >= 0) {
    state.overrides.splice(existingIndex, 1, nextOverride);
  } else {
    state.overrides.push(nextOverride);
  }

  state.overrides.sort((left, right) => left.date.localeCompare(right.date));
  elements.overrideLabel.value = "";

  persistState();
  renderEverything();
}

function renderEverything() {
  const schedule = buildSchedule();
  currentPlan = buildViewModel(schedule);
  renderHeadings(currentPlan);
  renderPlanRows(currentPlan.rows);
  renderStats(currentPlan.meta);
  renderOverrideList();
  applySizing(currentPlan.meta.totalDays);
}

function buildViewModel(schedule) {
  const tradeValue = (state.tradeValue || DEFAULT_STATE.tradeValue).trim();
  return {
    tradeValue,
    monthValue: state.monthValue,
    tradeHeadingText: `TRADE : ${tradeValue}`,
    monthHeadingText: formatMonthHeading(state.monthValue),
    rows: schedule.rows,
    meta: schedule.meta,
  };
}

function renderHeadings(viewModel) {
  elements.tradeHeadingText.textContent = viewModel.tradeHeadingText;
  elements.monthHeadingText.textContent = viewModel.monthHeadingText;
  document.title = viewModel.monthHeadingText;
}

function renderPlanRows(rows) {
  elements.planBody.replaceChildren();

  rows.forEach((row) => {
    const tr = document.createElement("tr");
    tr.append(createCell(row.dateLabel, "date-cell"));

    if (row.type === "work") {
      tr.append(createCell(row.jobNo));
      tr.append(createCell(row.professionalSkill));
      tr.append(createCell(row.lessonNo));
      tr.append(createCell(row.professionalKnowledge));
    } else if (row.type === "blank") {
      tr.append(createCell(""));
      tr.append(createCell(""));
      tr.append(createCell(""));
      tr.append(createCell(""));
    } else {
      tr.append(createCell(row.label, `merged-cell merged-cell--${row.type}`, 4));
    }

    elements.planBody.append(tr);
  });
}

function renderStats(meta) {
  elements.statsPanel.replaceChildren();

  const chips = [
    { text: `${meta.totalDays} calendar days` },
    { text: `${meta.weekendDays} weekend rows` },
    { text: `${meta.overrideCount} custom day rows` },
    { text: `${meta.importedItems} imported lesson rows` },
    { text: `${meta.scheduledItems} lesson rows placed` },
    { text: `${meta.examRows} exam rows` },
    { text: "Using Excel paste" },
  ];

  if (meta.unscheduledItems > 0) {
    chips.push({
      text: `${meta.unscheduledItems} lesson rows left outside the month`,
      className: "stat-chip--warning",
    });
  }

  if (meta.unplacedExamGroup) {
    chips.push({
      text: `EXAM LO ${meta.unplacedExamGroup} could not be placed`,
      className: "stat-chip--warning",
    });
  }

  if (meta.importedItems === 0) {
    chips.push({
      text: "Paste Excel data to fill working days",
      className: "stat-chip--warning",
    });
  }

  chips.forEach((chip) => {
    const item = document.createElement("div");
    item.className = ["stat-chip", chip.className].filter(Boolean).join(" ");
    item.textContent = chip.text;
    elements.statsPanel.append(item);
  });
}

function renderOverrideList() {
  const monthPrefix = `${state.monthValue}-`;
  const currentMonthOverrides = state.overrides.filter((override) => override.date.startsWith(monthPrefix));
  elements.overrideList.replaceChildren();

  if (!currentMonthOverrides.length) {
    const empty = document.createElement("div");
    empty.className = "stat-chip";
    empty.textContent = "No custom overrides added for this month.";
    elements.overrideList.append(empty);
    return;
  }

  currentMonthOverrides.forEach((override) => {
    const chip = document.createElement("div");
    chip.className = "override-chip";

    const typeLabel = {
      leave: "Leave",
      "note-blue": "Blue",
      "note-dark": "Dark blue",
    }[override.type];

    const label = document.createElement("span");
    label.textContent = `${formatDisplayDate(parseIsoDate(override.date))} | ${typeLabel} | ${override.label}`;

    const removeButton = document.createElement("button");
    removeButton.type = "button";
    removeButton.dataset.date = override.date;
    removeButton.setAttribute("aria-label", `Remove override for ${override.date}`);
    removeButton.textContent = "x";

    chip.append(label, removeButton);
    elements.overrideList.append(chip);
  });
}

function applySizing(totalDays) {
  const rowHeightMm = 209 / Math.max(totalDays, 1);
  const cellFontSize =
    totalDays >= 31 ? 11.35 :
    totalDays >= 30 ? 11.7 :
    totalDays >= 29 ? 12.3 : 13.1;
  const bandFontSize =
    totalDays >= 31 ? 12.9 :
    totalDays >= 30 ? 13.2 : 13.8;

  elements.planPage.style.setProperty("--body-row-height", `${rowHeightMm.toFixed(3)}mm`);
  elements.planPage.style.setProperty("--cell-font-size", `${cellFontSize}px`);
  elements.planPage.style.setProperty("--band-font-size", `${bandFontSize}px`);
}

async function downloadExcel() {
  renderEverything();

  if (!currentPlan) {
    return;
  }

  const available = await ensureExportService();
  if (!available) {
    window.alert(
      "Excel download needs the local export server.\n\nRun launch-monthly-plan.cmd in this folder, then try again.",
    );
    return;
  }

  const button = elements.downloadExcelButton;
  const idleLabel = button.textContent;
  button.disabled = true;
  button.textContent = "Preparing Excel...";
  setExportStatus("info", "Generating Excel file...");

  try {
    const response = await fetch(`${EXPORT_SERVICE_ORIGIN}/api/export/xlsx`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(currentPlan),
    });

    if (!response.ok) {
      const detail = await response.text();
      throw new Error(detail || `Export failed with status ${response.status}`);
    }

    const blob = await response.blob();
    const fileName = getFileNameFromResponse(response, currentPlan);
    downloadBlob(blob, fileName);
    setExportStatus("success", "Excel download started.");
  } catch (error) {
    setExportStatus("error", error.message || "Export failed.");
  } finally {
    button.disabled = false;
    button.textContent = idleLabel;
  }
}

async function checkExportService() {
  try {
    const response = await fetch(`${EXPORT_SERVICE_ORIGIN}/health`, {
      method: "GET",
      cache: "no-store",
    });

    exportServiceReady = response.ok;
  } catch {
    exportServiceReady = false;
  }

  updateExportAvailability();
}

async function ensureExportService() {
  if (exportServiceReady) {
    return true;
  }

  await checkExportService();
  return exportServiceReady;
}

function updateExportAvailability() {
  elements.downloadExcelButton.disabled = !exportServiceReady;
}

function setExportStatus(type, message) {
  elements.exportStatus.className = `export-status ${type === "error" ? "is-error" : ""} ${type === "success" ? "is-success" : ""}`.trim();
  elements.exportStatus.textContent = message;
}

function clearExportStatus() {
  elements.exportStatus.className = "export-status";
  elements.exportStatus.textContent = "";
}

function getFileNameFromResponse(response, viewModel) {
  const disposition = response.headers.get("Content-Disposition") || "";
  const match = disposition.match(/filename="?([^"]+)"?/i);
  if (match) {
    return match[1];
  }

  const monthPart = (viewModel.monthValue || "monthly-plan").replace(/[^0-9-]/g, "");
  return `monthly-plan-${monthPart}.xlsx`;
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.append(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 2000);
}

function buildSchedule() {
  const dates = getMonthDates(state.monthValue);
  const overrideMap = new Map(
    state.overrides
      .filter((override) => override.date.startsWith(`${state.monthValue}-`))
      .map((override) => [override.date, override]),
  );

  const imported = getImportedItems();
  const rows = [];
  let dataIndex = 0;
  let pendingExamGroup = "";
  let examRows = 0;
  let weekendDays = 0;

  dates.forEach((date) => {
    const isoDate = formatIsoDate(date);
    const override = overrideMap.get(isoDate);
    const weekendLabel = getWeekendLabel(date);

    if (override) {
      rows.push({
        type: override.type,
        dateLabel: formatDisplayDate(date),
        label: override.type === "leave" ? override.label.toUpperCase() : override.label,
      });
      return;
    }

    if (weekendLabel) {
      weekendDays += 1;
      rows.push({
        type: "leave",
        dateLabel: formatDisplayDate(date),
        label: weekendLabel,
      });
      return;
    }

    if (pendingExamGroup) {
      examRows += 1;
      rows.push({
        type: "exam",
        dateLabel: formatDisplayDate(date),
        label: `EXAM LO ${pendingExamGroup}`,
      });
      pendingExamGroup = "";
      return;
    }

    const item = imported.items[dataIndex];

    if (!item) {
      rows.push({
        type: "blank",
        dateLabel: formatDisplayDate(date),
      });
      return;
    }

    rows.push({
      type: "work",
      dateLabel: formatDisplayDate(date),
      jobNo: item.jobNo,
      professionalSkill: item.professionalSkill,
      lessonNo: item.lessonNo,
      professionalKnowledge: item.professionalKnowledge,
    });

    const currentGroup = extractJobGroup(item.jobNo);
    const nextItem = imported.items[dataIndex + 1];
    const nextGroup = nextItem ? extractJobGroup(nextItem.jobNo) : "";

    if (currentGroup && nextGroup && currentGroup !== nextGroup) {
      pendingExamGroup = currentGroup;
    }

    dataIndex += 1;
  });

  return {
    rows,
    meta: {
      totalDays: dates.length,
      weekendDays,
      overrideCount: overrideMap.size,
      importedItems: imported.items.length,
      scheduledItems: dataIndex,
      unscheduledItems: Math.max(imported.items.length - dataIndex, 0),
      examRows,
      unplacedExamGroup: pendingExamGroup,
      dataSource: imported.source,
    },
  };
}

function getImportedItems() {
  const skillRows = parseTwoColumnRows(state.skillData, "JOB NO", "PROFESSIONAL SKILL");
  const knowledgeRows = parseTwoColumnRows(state.knowledgeData, "LESSON NO", "PROFESSIONAL KNOWLEDGE");
  const length = Math.max(skillRows.length, knowledgeRows.length);

  const items = Array.from({ length }, (_, index) => ({
    jobNo: skillRows[index]?.left ?? "",
    professionalSkill: skillRows[index]?.right ?? "",
    lessonNo: knowledgeRows[index]?.left ?? "",
    professionalKnowledge: knowledgeRows[index]?.right ?? "",
  })).filter((row) => Object.values(row).some((value) => value.trim() !== ""));

  return { source: "split", items };
}

function parseTwoColumnRows(text, firstHeader, secondHeader) {
  const lines = toUsefulLines(text);
  const rows = [];

  lines.forEach((line) => {
    const parts = splitColumns(line);
    if (parts.length < 2) {
      return;
    }

    const left = cleanCell(parts[0]);
    const right = cleanCell(parts.slice(1).join(" "));

    if (looksLikeHeader(`${left} ${right}`, firstHeader, secondHeader)) {
      return;
    }

    rows.push({ left, right });
  });

  return rows;
}

function splitColumns(line) {
  if (line.includes("\t")) {
    return line.split("\t");
  }

  return line.split(/\s{2,}/);
}

function looksLikeHeader(text, firstHeader, secondHeader) {
  const normalized = text.replace(/\./g, "").replace(/\s+/g, " ").trim().toUpperCase();
  return normalized.includes(firstHeader.replace(/\./g, "").toUpperCase())
    && normalized.includes(secondHeader.replace(/\./g, "").toUpperCase());
}

function toUsefulLines(text) {
  return text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
}

function cleanCell(value) {
  return value.replace(/^"(.*)"$/, "$1").replace(/\s+/g, " ").trim();
}

function getMonthDates(monthValue) {
  const [yearText, monthText] = monthValue.split("-");
  const year = Number(yearText);
  const monthIndex = Number(monthText) - 1;

  if (!Number.isFinite(year) || !Number.isFinite(monthIndex)) {
    return [];
  }

  const dayCount = new Date(Date.UTC(year, monthIndex + 1, 0)).getUTCDate();
  return Array.from({ length: dayCount }, (_, index) => new Date(Date.UTC(year, monthIndex, index + 1)));
}

function formatMonthHeading(monthValue) {
  const dates = getMonthDates(monthValue);
  if (!dates.length) {
    return "MONTHLY PLAN";
  }

  const date = dates[0];
  const monthName = date.toLocaleString("en-US", { month: "long", timeZone: "UTC" }).toUpperCase();
  const year = date.getUTCFullYear();
  return `MONTHLY PLAN - ${monthName} -${year}`;
}

function formatDisplayDate(date) {
  const day = String(date.getUTCDate()).padStart(2, "0");
  const month = String(date.getUTCMonth() + 1).padStart(2, "0");
  const year = date.getUTCFullYear();
  return `${day}-${month}-${year}`;
}

function formatIsoDate(date) {
  return [
    date.getUTCFullYear(),
    String(date.getUTCMonth() + 1).padStart(2, "0"),
    String(date.getUTCDate()).padStart(2, "0"),
  ].join("-");
}

function parseIsoDate(value) {
  const [year, month, day] = value.split("-").map(Number);
  return new Date(Date.UTC(year, month - 1, day));
}

function getWeekendLabel(date) {
  const day = date.getUTCDay();
  if (day === 0) {
    return "SUNDAY";
  }
  if (day === 6) {
    return "SATURDAY";
  }
  return "";
}

function extractJobGroup(jobNo) {
  const match = String(jobNo || "").trim().match(/^([0-9]+)/);
  return match ? match[1] : "";
}

function createCell(text, className = "", colSpan = 1) {
  const td = document.createElement("td");
  td.textContent = text || "";
  if (className) {
    td.className = className;
  }
  if (colSpan > 1) {
    td.colSpan = colSpan;
  }
  return td;
}

function persistState() {
  window.localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function loadState() {
  try {
    const saved = JSON.parse(window.localStorage.getItem(STORAGE_KEY) || "null");
    if (!saved || typeof saved !== "object") {
      return structuredClone(DEFAULT_STATE);
    }

    return {
      ...structuredClone(DEFAULT_STATE),
      ...saved,
      overrides: Array.isArray(saved.overrides) ? saved.overrides : [],
    };
  } catch {
    return structuredClone(DEFAULT_STATE);
  }
}
