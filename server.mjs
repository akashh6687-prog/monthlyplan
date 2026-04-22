import http from "node:http";
import os from "node:os";
import path from "node:path";
import fs from "node:fs/promises";
import { access } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { chromium } from "playwright-core";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const HOST = "127.0.0.1";
const PORT = 3210;
const MAX_BODY_BYTES = 2 * 1024 * 1024;
const TABLE_LINE = "#303030";
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const STYLESHEET = await fs.readFile(path.join(__dirname, "styles.css"), "utf8");

const BROWSER_PATHS = [
  "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
  "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe",
];

function setCorsHeaders(res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  res.setHeader("Access-Control-Expose-Headers", "Content-Disposition");
}

function sendJson(res, statusCode, payload) {
  setCorsHeaders(res);
  res.writeHead(statusCode, { "Content-Type": "application/json; charset=utf-8" });
  res.end(JSON.stringify(payload));
}

function sendText(res, statusCode, message) {
  setCorsHeaders(res);
  res.writeHead(statusCode, { "Content-Type": "text/plain; charset=utf-8" });
  res.end(message);
}

async function readJsonBody(req) {
  const chunks = [];
  let total = 0;

  for await (const chunk of req) {
    total += chunk.length;
    if (total > MAX_BODY_BYTES) {
      throw new Error("Request body is too large.");
    }
    chunks.push(chunk);
  }

  if (!chunks.length) {
    return {};
  }

  return JSON.parse(Buffer.concat(chunks).toString("utf8"));
}

function sanitizeText(value, fallback = "") {
  return String(value ?? fallback)
    .replace(/\r/g, "")
    .trim();
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function sanitizeFilename(value) {
  return String(value)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 80) || "monthly-plan";
}

function normalizeRows(rows) {
  const allowedTypes = new Set(["work", "blank", "leave", "exam", "note-blue", "note-dark"]);

  return (Array.isArray(rows) ? rows : []).slice(0, 31).map((row) => {
    const type = allowedTypes.has(row?.type) ? row.type : "blank";
    return {
      type,
      dateLabel: sanitizeText(row?.dateLabel),
      jobNo: sanitizeText(row?.jobNo),
      professionalSkill: sanitizeText(row?.professionalSkill),
      lessonNo: sanitizeText(row?.lessonNo),
      professionalKnowledge: sanitizeText(row?.professionalKnowledge),
      label: sanitizeText(row?.label),
    };
  });
}

function normalizePlanPayload(payload) {
  const rows = normalizeRows(payload?.rows);
  const totalDays = rows.length;

  return {
    tradeHeadingText: sanitizeText(payload?.tradeHeadingText, "TRADE :"),
    monthHeadingText: sanitizeText(payload?.monthHeadingText, "MONTHLY PLAN"),
    tradeValue: sanitizeText(payload?.tradeValue),
    monthValue: sanitizeText(payload?.monthValue),
    rows,
    meta: {
      totalDays,
    },
  };
}

function computeSizing(totalDays) {
  const bodyRowHeightMm = 209 / Math.max(totalDays, 1);
  const cellFontSize =
    totalDays >= 31 ? 11.35 :
    totalDays >= 30 ? 11.7 :
    totalDays >= 29 ? 12.3 : 13.1;
  const bandFontSize =
    totalDays >= 31 ? 12.9 :
    totalDays >= 30 ? 13.2 : 13.8;

  return {
    bodyRowHeightMm,
    bodyRowHeightPx: Math.max(23, Math.round(bodyRowHeightMm * 3.7795275591)),
    cellFontSize,
    bandFontSize,
  };
}

function buildPlanRowsHtml(rows) {
  return rows.map((row) => {
    const dateCell = `<td class="date-cell">${escapeHtml(row.dateLabel)}</td>`;

    if (row.type === "work") {
      return `
        <tr>
          ${dateCell}
          <td>${escapeHtml(row.jobNo)}</td>
          <td>${escapeHtml(row.professionalSkill)}</td>
          <td>${escapeHtml(row.lessonNo)}</td>
          <td>${escapeHtml(row.professionalKnowledge)}</td>
        </tr>
      `;
    }

    if (row.type === "blank") {
      return `
        <tr>
          ${dateCell}
          <td></td>
          <td></td>
          <td></td>
          <td></td>
        </tr>
      `;
    }

    return `
      <tr>
        ${dateCell}
        <td class="merged-cell merged-cell--${escapeHtml(row.type)}" colspan="4">${escapeHtml(row.label)}</td>
      </tr>
    `;
  }).join("");
}

function buildPrintDocument(model) {
  const sizing = computeSizing(model.meta.totalDays);
  const styleVars = [
    `--body-row-height:${sizing.bodyRowHeightMm.toFixed(3)}mm`,
    `--cell-font-size:${sizing.cellFontSize}px`,
    `--band-font-size:${sizing.bandFontSize}px`,
  ].join(";");

  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>${escapeHtml(model.monthHeadingText)}</title>
  <style>
${STYLESHEET}
html, body {
  margin: 0;
  padding: 0;
  background: white !important;
}
body {
  min-height: auto;
}
.app-shell,
.controls-pane,
.preview-header {
  display: none !important;
}
.export-wrap {
  padding: 0;
  margin: 0;
}
.plan-page {
  margin: 0;
  box-shadow: none;
}
  </style>
</head>
<body>
  <div class="export-wrap">
    <section class="plan-page" style="${styleVars}">
      <div class="plan-sheet">
        <div class="plan-band plan-band--trade">
          <span>${escapeHtml(model.tradeHeadingText)}</span>
        </div>
        <div class="plan-band plan-band--month">
          <span>${escapeHtml(model.monthHeadingText)}</span>
        </div>
        <table class="plan-table" aria-label="Monthly plan export">
          <colgroup>
            <col class="col-date">
            <col class="col-job">
            <col class="col-skill">
            <col class="col-lesson">
            <col class="col-knowledge">
          </colgroup>
          <thead>
            <tr>
              <th scope="col">DATE</th>
              <th scope="col">JOB<br>NO.</th>
              <th scope="col">PROFESSIONAL SKILL</th>
              <th scope="col">LESSON<br>NO.</th>
              <th scope="col">PROFESSIONAL KNOWLEDGE</th>
            </tr>
          </thead>
          <tbody>
            ${buildPlanRowsHtml(model.rows)}
          </tbody>
        </table>
        <footer class="signature-strip">
          <span>INSTRUCTOR</span>
          <span>GROUP INSTRUCTOR</span>
          <span>PRINCIPAL</span>
        </footer>
      </div>
    </section>
  </div>
</body>
</html>`;
}

function applySingleCellBorder(range) {
  range.format.borders = { preset: "outside", style: "thin", color: TABLE_LINE };
}

function formatNormalBodyCell(range, alignment = "center") {
  range.format = {
    horizontalAlignment: alignment,
    verticalAlignment: "center",
    wrapText: true,
    font: { name: "Arial", size: 10 },
  };
  applySingleCellBorder(range);
}

export async function buildWorkbookFromPlan(payload) {
  const model = normalizePlanPayload(payload);
  const sizing = computeSizing(model.meta.totalDays);
  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add("Monthly Plan");
  sheet.showGridLines = false;

  sheet.getRange("A1:A80").format.columnWidthPx = 92;
  sheet.getRange("B1:B80").format.columnWidthPx = 72;
  sheet.getRange("C1:C80").format.columnWidthPx = 268;
  sheet.getRange("D1:D80").format.columnWidthPx = 84;
  sheet.getRange("E1:E80").format.columnWidthPx = 280;

  const tableStartRow = 4;
  const lastDataRow = tableStartRow + model.rows.length - 1;
  const signatureTopRow = lastDataRow + 1;
  const signatureLabelRow = signatureTopRow + 4;
  const sheetEndRow = signatureLabelRow;

  sheet.getRange(`A1:E${sheetEndRow}`).format = {
    font: { name: "Arial", size: 10, color: "#111111" },
    verticalAlignment: "center",
  };
  sheet.getRange(`A1:E${sheetEndRow}`).format.borders = {
    preset: "outside",
    style: "thin",
    color: TABLE_LINE,
  };

  const tradeRange = sheet.getRange("A1:E1");
  tradeRange.merge();
  tradeRange.values = [[model.tradeHeadingText]];
  tradeRange.format = {
    fill: "#dce8f5",
    font: { name: "Arial", size: 12, bold: true, color: "#12356c" },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
  };
  applySingleCellBorder(tradeRange);
  tradeRange.format.rowHeightPx = 31;

  const monthRange = sheet.getRange("A2:E2");
  monthRange.merge();
  monthRange.values = [[model.monthHeadingText]];
  monthRange.format = {
    fill: "#f7c9a7",
    font: { name: "Arial", size: 12, bold: true, color: "#111111" },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
  };
  applySingleCellBorder(monthRange);
  monthRange.format.rowHeightPx = 38;

  const headerRange = sheet.getRange("A3:E3");
  headerRange.values = [[
    "DATE",
    "JOB\nNO.",
    "PROFESSIONAL SKILL",
    "LESSON\nNO.",
    "PROFESSIONAL KNOWLEDGE",
  ]];
  headerRange.format = {
    fill: "#d8e7d0",
    font: { name: "Arial", size: 11, bold: true, color: "#111111" },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
  };
  headerRange.format.rowHeightPx = 36;
  ["A3", "B3", "C3", "D3", "E3"].forEach((address) => applySingleCellBorder(sheet.getRange(address)));

  model.rows.forEach((row, index) => {
    const excelRow = tableStartRow + index;
    const rowRange = sheet.getRange(`A${excelRow}:E${excelRow}`);
    rowRange.format.rowHeightPx = sizing.bodyRowHeightPx;

    const dateCell = sheet.getRange(`A${excelRow}`);
    dateCell.values = [[row.dateLabel]];
    dateCell.format = {
      horizontalAlignment: "center",
      verticalAlignment: "center",
      wrapText: false,
      font: { name: "Arial", size: 10, color: "#111111" },
    };
    applySingleCellBorder(dateCell);

    if (row.type === "work") {
      const values = [row.jobNo, row.professionalSkill, row.lessonNo, row.professionalKnowledge];
      ["B", "C", "D", "E"].forEach((column, columnIndex) => {
        const cell = sheet.getRange(`${column}${excelRow}`);
        cell.values = [[values[columnIndex]]];
      });

      formatNormalBodyCell(sheet.getRange(`B${excelRow}`));
      sheet.getRange(`B${excelRow}`).format.font = { name: "Arial", size: 10, bold: true };

      formatNormalBodyCell(sheet.getRange(`C${excelRow}`), "left");

      formatNormalBodyCell(sheet.getRange(`D${excelRow}`));
      formatNormalBodyCell(sheet.getRange(`E${excelRow}`), "left");
      return;
    }

    if (row.type === "blank") {
      ["B", "C", "D", "E"].forEach((column) => {
        const cell = sheet.getRange(`${column}${excelRow}`);
        cell.values = [[""]];
        formatNormalBodyCell(cell, column === "E" ? "left" : "center");
      });
      return;
    }

    const mergedRange = sheet.getRange(`B${excelRow}:E${excelRow}`);
    mergedRange.merge();
    mergedRange.values = [[row.label]];
    mergedRange.format = {
      fill:
        row.type === "leave" ? "#ff9fa4" :
        row.type === "note-dark" ? "#92b6db" : "#b7c9ea",
      font: { name: "Arial", size: 11, bold: true, color: "#111111" },
      horizontalAlignment: "center",
      verticalAlignment: "center",
      wrapText: true,
    };
    applySingleCellBorder(mergedRange);
  });

  sheet.getRange(`A${signatureTopRow}:E${signatureLabelRow}`).format.rowHeightPx = 28;
  sheet.getRange(`A${signatureTopRow}:E${signatureLabelRow}`).format = {
    horizontalAlignment: "center",
    verticalAlignment: "center",
    font: { name: "Arial", size: 10, bold: true, color: "#111111" },
  };

  sheet.getRange(`A${signatureLabelRow}`).values = [["INSTRUCTOR"]];
  sheet.getRange(`C${signatureLabelRow}`).values = [["GROUP INSTRUCTOR"]];
  sheet.getRange(`E${signatureLabelRow}`).values = [["PRINCIPAL"]];

  return workbook;
}

async function writeWorkbookToTempXlsx(workbook) {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "iti-monthly-plan-"));
  const outputPath = path.join(tempDir, "monthly-plan.xlsx");
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(outputPath);
  const buffer = await fs.readFile(outputPath);
  await fs.rm(tempDir, { recursive: true, force: true });
  return buffer;
}

async function detectBrowserPath() {
  for (const browserPath of BROWSER_PATHS) {
    try {
      await access(browserPath);
      return browserPath;
    } catch {
      continue;
    }
  }

  throw new Error("No supported browser was found for PDF export.");
}

export async function exportXlsxBufferFromPlan(payload) {
  const workbook = await buildWorkbookFromPlan(payload);
  return writeWorkbookToTempXlsx(workbook);
}

export async function renderPdfBufferFromPlan(payload) {
  const model = normalizePlanPayload(payload);
  const browserPath = await detectBrowserPath();
  const browser = await chromium.launch({
    headless: true,
    executablePath: browserPath,
  });

  try {
    const page = await browser.newPage({
      viewport: { width: 1200, height: 1700 },
      deviceScaleFactor: 1,
    });
    await page.setContent(buildPrintDocument(model), { waitUntil: "load" });
    return await page.pdf({
      format: "A4",
      printBackground: true,
      preferCSSPageSize: true,
      margin: { top: "0", right: "0", bottom: "0", left: "0" },
    });
  } finally {
    await browser.close();
  }
}

function getDownloadName(model, extension) {
  const base =
    model.monthValue
      ? `monthly-plan-${model.monthValue}`
      : sanitizeFilename(model.monthHeadingText || "monthly-plan");
  return `${base}.${extension}`;
}

function createServer() {
  return http.createServer(async (req, res) => {
    setCorsHeaders(res);

    if (req.method === "OPTIONS") {
      res.writeHead(204);
      res.end();
      return;
    }

    if (req.method === "GET" && req.url === "/health") {
      sendJson(res, 200, { ok: true, service: "monthly-plan-export" });
      return;
    }

    if (req.method === "POST" && req.url === "/api/export/xlsx") {
      try {
        const payload = await readJsonBody(req);
        const model = normalizePlanPayload(payload);
        const buffer = await exportXlsxBufferFromPlan(model);
        res.writeHead(200, {
          "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          "Content-Disposition": `attachment; filename="${getDownloadName(model, "xlsx")}"`,
        });
        res.end(buffer);
      } catch (error) {
        sendText(res, 500, error.message || "Excel export failed.");
      }
      return;
    }

    if (req.method === "POST" && req.url === "/api/export/pdf") {
      try {
        const payload = await readJsonBody(req);
        const model = normalizePlanPayload(payload);
        const buffer = await renderPdfBufferFromPlan(model);
        res.writeHead(200, {
          "Content-Type": "application/pdf",
          "Content-Disposition": `attachment; filename="${getDownloadName(model, "pdf")}"`,
        });
        res.end(buffer);
      } catch (error) {
        sendText(res, 500, error.message || "PDF export failed.");
      }
      return;
    }

    sendText(res, 404, "Not found.");
  });
}

export function startServer({ host = HOST, port = PORT } = {}) {
  const server = createServer();
  return new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(port, host, () => {
      server.off("error", reject);
      console.log(`Monthly plan export server running on http://${host}:${port}`);
      resolve(server);
    });
  });
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  startServer().catch((error) => {
    console.error(error.message || error);
    process.exitCode = 1;
  });
}
