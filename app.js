const STORAGE_KEY = "iti-monthly-plan-maker-v2";
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
  downloadExcelButton: document.getElementById("downloadExcelButton"),
  printButton: document.getElementById("printButton"),
  resetButton: document.getElementById("resetButton"),
  statsPanel: document.getElementById("statsPanel"),
  tradeHeadingText: document.getElementById("tradeHeadingText"),
  monthHeadingText: document.getElementById("monthHeadingText"),
  planBody: document.getElementById("planBody"),
  planPage: document.getElementById("planPage"),
  exportStatus: document.getElementById("exportStatus"),
};

let state = loadState();
let currentPlan = null;
const CRC32_TABLE = createCrc32Table();
const XLSX_STYLE_IDS = {
  trade: 1,
  month: 2,
  header: 3,
  date: 4,
  job: 5,
  skill: 6,
  lesson: 7,
  knowledge: 8,
  leave: 9,
  exam: 10,
  noteBlue: 11,
  noteDark: 12,
  blank: 13,
  signature: 14,
};

bootstrap();

function bootstrap() {
  syncFormWithState();
  bindEvents();
  syncOverrideDateOptions();
  renderEverything();
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
  elements.downloadExcelButton.addEventListener("click", downloadExcelWorkbook);
  elements.printButton.addEventListener("click", () => {
    renderEverything();
    setStatus("Browser print opened with print-color styling enabled.");
    window.print();
  });

  elements.resetButton.addEventListener("click", () => {
    const confirmed = window.confirm("Reset all inputs and remove saved data for this planner?");
    if (!confirmed) {
      return;
    }

    state = structuredClone(DEFAULT_STATE);
    persistState();
    syncFormWithState();
    syncOverrideDateOptions();
    clearStatus();
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

function setStatus(message, type = "") {
  elements.exportStatus.className = `export-status ${type === "error" ? "is-error" : ""} ${type === "success" ? "is-success" : ""}`.trim();
  elements.exportStatus.textContent = message;
}

function clearStatus() {
  elements.exportStatus.className = "export-status";
  elements.exportStatus.textContent = "";
}

function downloadExcelWorkbook() {
  renderEverything();

  try {
    const workbookBlob = buildMonthlyPlanWorkbookBlob(currentPlan);
    const fileName = `${sanitizeFileName(currentPlan?.monthHeadingText || "MONTHLY PLAN")}.xlsx`;
    downloadBlob(workbookBlob, fileName);
    setStatus(`Excel file downloaded as ${fileName}.`, "success");
  } catch (error) {
    console.error(error);
    setStatus("Excel download failed in the browser. Please try again.", "error");
  }
}

function buildMonthlyPlanWorkbookBlob(viewModel) {
  const files = [
    { name: "[Content_Types].xml", data: createContentTypesXml() },
    { name: "_rels/.rels", data: createRootRelationshipsXml() },
    { name: "xl/workbook.xml", data: createWorkbookXml() },
    { name: "xl/_rels/workbook.xml.rels", data: createWorkbookRelationshipsXml() },
    { name: "xl/styles.xml", data: createStylesXml() },
    { name: "xl/worksheets/sheet1.xml", data: createWorksheetXml(viewModel) },
  ];

  return createZipBlob(
    files,
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  );
}

function createWorksheetXml(viewModel) {
  const rowsXml = [];
  const mergeRefs = ["A1:E1", "A2:E2"];
  const totalDays = Math.max(viewModel?.meta?.totalDays || 0, 1);
  const bodyRowHeightPt = clampNumber((209 / totalDays) * 2.834645669, 17, 22);

  rowsXml.push(createRowXml(1, 23, [
    createInlineStringCell("A1", viewModel.tradeHeadingText, XLSX_STYLE_IDS.trade),
  ]));

  rowsXml.push(createRowXml(2, 27, [
    createInlineStringCell("A2", viewModel.monthHeadingText, XLSX_STYLE_IDS.month),
  ]));

  rowsXml.push(createRowXml(3, 26, [
    createInlineStringCell("A3", "DATE", XLSX_STYLE_IDS.header),
    createInlineStringCell("B3", "JOB\nNO.", XLSX_STYLE_IDS.header),
    createInlineStringCell("C3", "PROFESSIONAL SKILL", XLSX_STYLE_IDS.header),
    createInlineStringCell("D3", "LESSON\nNO.", XLSX_STYLE_IDS.header),
    createInlineStringCell("E3", "PROFESSIONAL KNOWLEDGE", XLSX_STYLE_IDS.header),
  ]));

  let rowIndex = 4;

  viewModel.rows.forEach((row) => {
    if (row.type === "work") {
      rowsXml.push(createRowXml(rowIndex, bodyRowHeightPt, [
        createInlineStringCell(`A${rowIndex}`, row.dateLabel, XLSX_STYLE_IDS.date),
        createInlineStringCell(`B${rowIndex}`, row.jobNo, XLSX_STYLE_IDS.job),
        createInlineStringCell(`C${rowIndex}`, row.professionalSkill, XLSX_STYLE_IDS.skill),
        createInlineStringCell(`D${rowIndex}`, row.lessonNo, XLSX_STYLE_IDS.lesson),
        createInlineStringCell(`E${rowIndex}`, row.professionalKnowledge, XLSX_STYLE_IDS.knowledge),
      ]));
    } else if (row.type === "blank") {
      rowsXml.push(createRowXml(rowIndex, bodyRowHeightPt, [
        createInlineStringCell(`A${rowIndex}`, row.dateLabel, XLSX_STYLE_IDS.date),
        createEmptyCell(`B${rowIndex}`, XLSX_STYLE_IDS.blank),
        createEmptyCell(`C${rowIndex}`, XLSX_STYLE_IDS.blank),
        createEmptyCell(`D${rowIndex}`, XLSX_STYLE_IDS.blank),
        createEmptyCell(`E${rowIndex}`, XLSX_STYLE_IDS.blank),
      ]));
    } else {
      mergeRefs.push(`B${rowIndex}:E${rowIndex}`);
      rowsXml.push(createRowXml(rowIndex, bodyRowHeightPt, [
        createInlineStringCell(`A${rowIndex}`, row.dateLabel, XLSX_STYLE_IDS.date),
        createInlineStringCell(`B${rowIndex}`, row.label, getMergedRowStyleId(row.type)),
      ]));
    }

    rowIndex += 1;
  });

  rowsXml.push(createRowXml(rowIndex, 14, [
    createEmptyCell(`A${rowIndex}`, XLSX_STYLE_IDS.blank),
    createEmptyCell(`B${rowIndex}`, XLSX_STYLE_IDS.blank),
    createEmptyCell(`C${rowIndex}`, XLSX_STYLE_IDS.blank),
    createEmptyCell(`D${rowIndex}`, XLSX_STYLE_IDS.blank),
    createEmptyCell(`E${rowIndex}`, XLSX_STYLE_IDS.blank),
  ]));
  rowIndex += 1;

  mergeRefs.push(`A${rowIndex}:B${rowIndex}`, `C${rowIndex}:D${rowIndex}`);
  rowsXml.push(createRowXml(rowIndex, 26, [
    createInlineStringCell(`A${rowIndex}`, "INSTRUCTOR", XLSX_STYLE_IDS.signature),
    createInlineStringCell(`C${rowIndex}`, "GROUP INSTRUCTOR", XLSX_STYLE_IDS.signature),
    createInlineStringCell(`E${rowIndex}`, "PRINCIPAL", XLSX_STYLE_IDS.signature),
  ]));

  const lastRow = rowIndex;
  const mergesXml = mergeRefs.length
    ? `<mergeCells count="${mergeRefs.length}">${mergeRefs.map((ref) => `<mergeCell ref="${ref}"/>`).join("")}</mergeCells>`
    : "";

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:E${lastRow}"/>
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/>
      <selection pane="bottomLeft" activeCell="A4" sqref="A4"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols>
    <col min="1" max="1" width="12.5" customWidth="1"/>
    <col min="2" max="2" width="9.5" customWidth="1"/>
    <col min="3" max="3" width="34" customWidth="1"/>
    <col min="4" max="4" width="10.5" customWidth="1"/>
    <col min="5" max="5" width="36" customWidth="1"/>
  </cols>
  <sheetData>${rowsXml.join("")}</sheetData>
  ${mergesXml}
  <pageMargins left="0.25" right="0.25" top="0.35" bottom="0.35" header="0.2" footer="0.2"/>
  <pageSetup paperSize="9" orientation="portrait" fitToWidth="1" fitToHeight="0"/>
</worksheet>`;
}

function createContentTypesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`;
}

function createRootRelationshipsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
}

function createWorkbookXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <bookViews>
    <workbookView xWindow="0" yWindow="0" windowWidth="24000" windowHeight="12840"/>
  </bookViews>
  <sheets>
    <sheet name="Monthly Plan" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`;
}

function createWorkbookRelationshipsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
}

function createStylesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font>
      <sz val="11"/>
      <name val="Arial"/>
      <family val="2"/>
    </font>
    <font>
      <b/>
      <sz val="11"/>
      <name val="Arial"/>
      <family val="2"/>
    </font>
  </fonts>
  <fills count="8">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFDCE8F5"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF7C9A7"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFD8E7D0"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFF9FA4"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFB7C9EA"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF92B6DB"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left style="thin"><color rgb="FF303030"/></left>
      <right style="thin"><color rgb="FF303030"/></right>
      <top style="thin"><color rgb="FF303030"/></top>
      <bottom style="thin"><color rgb="FF303030"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="15">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="1" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="1" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="7" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>`;
}

function createZipBlob(files, mimeType) {
  const encoder = new TextEncoder();
  const zipParts = [];
  const centralDirectoryParts = [];
  let offset = 0;

  files.forEach((file) => {
    const nameBytes = encoder.encode(file.name);
    const dataBytes = typeof file.data === "string" ? encoder.encode(file.data) : file.data;
    const checksum = crc32(dataBytes);
    const { dosDate, dosTime } = getZipDateParts(new Date());

    const localHeader = new Uint8Array(30 + nameBytes.length);
    const localView = new DataView(localHeader.buffer);
    localView.setUint32(0, 0x04034b50, true);
    localView.setUint16(4, 20, true);
    localView.setUint16(6, 0, true);
    localView.setUint16(8, 0, true);
    localView.setUint16(10, dosTime, true);
    localView.setUint16(12, dosDate, true);
    localView.setUint32(14, checksum, true);
    localView.setUint32(18, dataBytes.length, true);
    localView.setUint32(22, dataBytes.length, true);
    localView.setUint16(26, nameBytes.length, true);
    localView.setUint16(28, 0, true);
    localHeader.set(nameBytes, 30);

    const centralHeader = new Uint8Array(46 + nameBytes.length);
    const centralView = new DataView(centralHeader.buffer);
    centralView.setUint32(0, 0x02014b50, true);
    centralView.setUint16(4, 20, true);
    centralView.setUint16(6, 20, true);
    centralView.setUint16(8, 0, true);
    centralView.setUint16(10, 0, true);
    centralView.setUint16(12, dosTime, true);
    centralView.setUint16(14, dosDate, true);
    centralView.setUint32(16, checksum, true);
    centralView.setUint32(20, dataBytes.length, true);
    centralView.setUint32(24, dataBytes.length, true);
    centralView.setUint16(28, nameBytes.length, true);
    centralView.setUint16(30, 0, true);
    centralView.setUint16(32, 0, true);
    centralView.setUint16(34, 0, true);
    centralView.setUint16(36, 0, true);
    centralView.setUint32(38, 0, true);
    centralView.setUint32(42, offset, true);
    centralHeader.set(nameBytes, 46);

    zipParts.push(localHeader, dataBytes);
    centralDirectoryParts.push(centralHeader);
    offset += localHeader.length + dataBytes.length;
  });

  const centralDirectorySize = centralDirectoryParts.reduce((total, part) => total + part.length, 0);
  const endRecord = new Uint8Array(22);
  const endView = new DataView(endRecord.buffer);
  endView.setUint32(0, 0x06054b50, true);
  endView.setUint16(4, 0, true);
  endView.setUint16(6, 0, true);
  endView.setUint16(8, files.length, true);
  endView.setUint16(10, files.length, true);
  endView.setUint32(12, centralDirectorySize, true);
  endView.setUint32(16, offset, true);
  endView.setUint16(20, 0, true);

  return new Blob([...zipParts, ...centralDirectoryParts, endRecord], { type: mimeType });
}

function createRowXml(rowNumber, heightPt, cells) {
  const height = Number.isFinite(heightPt) ? ` ht="${heightPt.toFixed(2)}" customHeight="1"` : "";
  return `<row r="${rowNumber}"${height}>${cells.join("")}</row>`;
}

function createInlineStringCell(cellRef, text, styleId) {
  return `<c r="${cellRef}" s="${styleId}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(text || "")}</t></is></c>`;
}

function createEmptyCell(cellRef, styleId) {
  return `<c r="${cellRef}" s="${styleId}"/>`;
}

function getMergedRowStyleId(type) {
  if (type === "leave") {
    return XLSX_STYLE_IDS.leave;
  }
  if (type === "exam") {
    return XLSX_STYLE_IDS.exam;
  }
  if (type === "note-dark") {
    return XLSX_STYLE_IDS.noteDark;
  }
  return XLSX_STYLE_IDS.noteBlue;
}

function sanitizeFileName(value) {
  const cleaned = String(value || "MONTHLY PLAN")
    .replace(/[<>:"/\\|?*\u0000-\u001f]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 120);

  return cleaned || "MONTHLY PLAN";
}

function downloadBlob(blob, fileName) {
  const link = document.createElement("a");
  const objectUrl = URL.createObjectURL(blob);

  link.href = objectUrl;
  link.download = fileName;
  link.style.display = "none";
  document.body.append(link);
  link.click();

  window.setTimeout(() => {
    link.remove();
    URL.revokeObjectURL(objectUrl);
  }, 1200);
}

function escapeXml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function clampNumber(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function getZipDateParts(date) {
  const safeYear = Math.max(date.getFullYear(), 1980);
  const dosTime =
    (date.getHours() << 11)
    | (date.getMinutes() << 5)
    | Math.floor(date.getSeconds() / 2);
  const dosDate =
    ((safeYear - 1980) << 9)
    | ((date.getMonth() + 1) << 5)
    | date.getDate();

  return { dosDate, dosTime };
}

function createCrc32Table() {
  const table = new Uint32Array(256);

  for (let index = 0; index < 256; index += 1) {
    let value = index;
    for (let bit = 0; bit < 8; bit += 1) {
      value = (value & 1) ? (0xedb88320 ^ (value >>> 1)) : (value >>> 1);
    }
    table[index] = value >>> 0;
  }

  return table;
}

function crc32(bytes) {
  let crc = 0xffffffff;

  for (let index = 0; index < bytes.length; index += 1) {
    crc = CRC32_TABLE[(crc ^ bytes[index]) & 0xff] ^ (crc >>> 8);
  }

  return (crc ^ 0xffffffff) >>> 0;
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
