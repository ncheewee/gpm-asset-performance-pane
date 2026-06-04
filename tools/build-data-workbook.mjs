import fs from "node:fs/promises";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const execFileAsync = promisify(execFile);

const here = dirname(fileURLToPath(import.meta.url));
const csvPath = join(here, "../data/portfolio-data.csv");
const outPath = join(here, "../data/portfolio-data.xlsx");
const tempDir = join(here, "../.tmp-workbook-build");
const tempXlsx = join(tempDir, "portfolio-data.raw.xlsx");
const unzipDir = join(tempDir, "xlsx");

const INPUT_FILL = "#FFF2CC";
const INPUT_FILL_SOFT = "#FFF8E8";
const OUTPUT_FILL = "#DDEBF7";
const OUTPUT_FILL_SOFT = "#EEF5FB";
const HEADER_FILL = "#1F4E78";
const HEADER_FONT = "#FFFFFF";
const BORDER = "#D9E2EC";

const inputFields = new Set([
  "name",
  "label",
  "country",
  "type",
  "gfa",
  "nla",
  "opex_m",
  "energy_gwh",
  "prev_cost_sqm",
  "prev_kwh_sqm",
  "trend_jun",
  "trend_jul",
  "trend_aug",
  "trend_sep",
  "trend_oct",
  "trend_nov",
  "trend_dec",
  "trend_jan",
  "trend_feb",
  "trend_mar",
  "trend_apr",
  "trend_may",
  "cost_energy_pct",
  "cost_cleaning_pct",
  "cost_security_pct",
  "cost_maintenance_pct",
  "cost_other_pct",
]);

const outputFields = new Set(["cost_sqm", "kwh_sqm", "score", "status"]);

function parseCsv(text) {
  return text.trim().split(/\r?\n/).map((line) => {
    const cells = [];
    let current = "";
    let quoted = false;
    for (let i = 0; i < line.length; i += 1) {
      const char = line[i];
      const next = line[i + 1];
      if (char === "\"" && quoted && next === "\"") {
        current += "\"";
        i += 1;
      } else if (char === "\"") {
        quoted = !quoted;
      } else if (char === "," && !quoted) {
        cells.push(current);
        current = "";
      } else {
        current += char;
      }
    }
    cells.push(current);
    return cells;
  });
}

function coerce(value) {
  if (value === "") return "";
  const numeric = Number(value);
  return Number.isFinite(numeric) && String(value).trim() !== "" ? numeric : value;
}

function columnLetter(index) {
  let letter = "";
  let n = index + 1;
  while (n > 0) {
    const remainder = (n - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    n = Math.floor((n - 1) / 26);
  }
  return letter;
}

function xmlEscape(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll("\"", "&quot;");
}

async function protectWorkbook(sheetName, inputCols, dataRowStart, dataRowEnd) {
  await fs.rm(unzipDir, { recursive: true, force: true });
  await fs.mkdir(unzipDir, { recursive: true });
  await execFileAsync("/usr/bin/unzip", ["-q", tempXlsx, "-d", unzipDir]);

  const workbookXmlPath = join(unzipDir, "xl/workbook.xml");
  const relsXmlPath = join(unzipDir, "xl/_rels/workbook.xml.rels");
  const workbookXml = await fs.readFile(workbookXmlPath, "utf8");
  const relsXml = await fs.readFile(relsXmlPath, "utf8");
  const escapedSheetName = xmlEscape(sheetName);
  const sheetMatch = workbookXml.match(new RegExp(`<(?:\\w+:)?sheet[^>]*name="${escapedSheetName}"[^>]*r:id="([^"]+)"[^>]*/>`));
  if (!sheetMatch) throw new Error(`Could not find worksheet "${sheetName}" in workbook.xml`);
  const relationshipId = sheetMatch[1];
  const relTag = [...relsXml.matchAll(/<Relationship\b[^>]*\/>/g)]
    .map((match) => match[0])
    .find((tag) => tag.includes(`Id="${relationshipId}"`));
  if (!relTag) throw new Error(`Could not resolve relationship "${relationshipId}"`);
  const targetMatch = relTag.match(/Target="([^"]+)"/);
  if (!targetMatch) throw new Error(`Could not read relationship target "${relationshipId}"`);
  const target = targetMatch[1].replace(/^\/xl\//, "");
  const sheetXmlPath = join(unzipDir, "xl", target);
  const stylesXmlPath = join(unzipDir, "xl/styles.xml");

  let stylesXml = await fs.readFile(stylesXmlPath, "utf8");
  const cellXfsMatch = stylesXml.match(/<((?:\w+:)?cellXfs) count="(\d+)">([\s\S]*?)<\/\1>/);
  if (!cellXfsMatch) throw new Error("Could not find cellXfs in styles.xml");

  const cellXfsTag = cellXfsMatch[1];
  const cellXfsBody = cellXfsMatch[3];
  const stylePrefixMatch = cellXfsTag.match(/^(\w+):/);
  const stylePrefix = stylePrefixMatch ? `${stylePrefixMatch[1]}:` : "";
  const xfPattern = new RegExp(`<${stylePrefix}xf\\b[^>]*/>|<${stylePrefix}xf\\b(?![^>]*\\/>)[^>]*>[\\s\\S]*?</${stylePrefix}xf>`, "g");
  const xfMatches = [...cellXfsBody.matchAll(xfPattern)].map((match) => match[0]);
  if (xfMatches.length !== Number(cellXfsMatch[2])) {
    throw new Error(`Expected ${cellXfsMatch[2]} cell styles, parsed ${xfMatches.length}`);
  }
  const unlockedStyleByOriginal = new Map();
  let unlockedXfs = "";

  function unlockedStyleIndex(originalIndex) {
    if (unlockedStyleByOriginal.has(originalIndex)) return unlockedStyleByOriginal.get(originalIndex);
    const originalXf = xfMatches[originalIndex] ?? xfMatches[0];
    let unlockedXf = originalXf
      .replace(/\sapplyProtection="[^"]*"/g, "")
      .replace(new RegExp(`<${stylePrefix}protection\\b[^>]*/>`, "g"), "")
      .replace(new RegExp(`<${stylePrefix}protection\\b[\\s\\S]*?</${stylePrefix}protection>`, "g"), "");
    if (unlockedXf.endsWith("/>")) {
      unlockedXf = unlockedXf.slice(0, -2) + ` applyProtection="1"><${stylePrefix}protection locked="0"/></${stylePrefix}xf>`;
    } else {
      unlockedXf = unlockedXf.replace(new RegExp(`</${stylePrefix}xf>$`), `<${stylePrefix}protection locked="0"/></${stylePrefix}xf>`);
      unlockedXf = unlockedXf.replace(new RegExp(`^<${stylePrefix}xf\\b([^>]*)>`), `<${stylePrefix}xf$1 applyProtection="1">`);
    }
    const newIndex = xfMatches.length + unlockedStyleByOriginal.size;
    unlockedStyleByOriginal.set(originalIndex, newIndex);
    unlockedXfs += unlockedXf;
    return newIndex;
  }

  let sheetXml = await fs.readFile(sheetXmlPath, "utf8");
  const inputColSet = new Set(inputCols);
  const worksheetTag = sheetXml.match(/<((?:\w+:)?worksheet)\b/)?.[1] ?? "worksheet";
  const worksheetPrefixMatch = worksheetTag.match(/^(\w+):/);
  const worksheetPrefix = worksheetPrefixMatch ? `${worksheetPrefixMatch[1]}:` : "";

  sheetXml = sheetXml.replace(new RegExp(`<${worksheetPrefix}c\\b([^>]*)\\br="([A-Z]+)(\\d+)"([^>]*)>`, "g"), (full, before, col, rowText, after) => {
    const row = Number(rowText);
    if (row < dataRowStart || row > dataRowEnd || !inputColSet.has(col)) return full;
    const styleMatch = full.match(/\bs="(\d+)"/);
    const originalStyle = styleMatch ? Number(styleMatch[1]) : 0;
    const newStyle = unlockedStyleIndex(originalStyle);
    if (styleMatch) return full.replace(/\bs="\d+"/, `s="${newStyle}"`);
    return `<${worksheetPrefix}c${before} r="${col}${rowText}"${after} s="${newStyle}">`;
  });

  const protection = `<${worksheetPrefix}sheetProtection sheet="1" objects="1" scenarios="1" selectLockedCells="0" selectUnlockedCells="1"/>`;
  if (!sheetXml.includes("<sheetProtection")) {
    sheetXml = sheetXml.replace(new RegExp(`(<${worksheetPrefix}sheetViews\\b[\\s\\S]*?</${worksheetPrefix}sheetViews>)`), `$1${protection}`);
  }

  if (unlockedStyleByOriginal.size > 0) {
    const newCount = xfMatches.length + unlockedStyleByOriginal.size;
    const newCellXfs = `<${cellXfsTag} count="${newCount}">${cellXfsBody}${unlockedXfs}</${cellXfsTag}>`;
    stylesXml = stylesXml.replace(cellXfsMatch[0], newCellXfs);
  }

  await fs.writeFile(sheetXmlPath, sheetXml);
  await fs.writeFile(stylesXmlPath, stylesXml);
  await fs.rm(outPath, { force: true });
  await execFileAsync("/usr/bin/zip", ["-qr", outPath, "."], { cwd: unzipDir });
  await fs.rm(tempDir, { recursive: true, force: true });
}

const csvText = await fs.readFile(csvPath, "utf8");
const [headers, ...rawRows] = parseCsv(csvText);
const rows = rawRows.map((row) => Object.fromEntries(headers.map((header, index) => [header, coerce(row[index] ?? "")])));
const workbook = Workbook.create();
const dataSheet = workbook.worksheets.add("Portfolio Data");
const guideSheet = workbook.worksheets.add("Field Guide");
dataSheet.showGridLines = false;
guideSheet.showGridLines = false;

const data = [headers, ...rows.map((row) => headers.map((header) => row[header]))];
dataSheet.getRange("A1").write(data);

const rowCount = rows.length;
const colCount = headers.length;
const lastCol = columnLetter(colCount - 1);
const inputCols = [];

dataSheet.freezePanes.freezeRows(1);
dataSheet.getRange(`A1:${lastCol}1`).format.fill.color = HEADER_FILL;
dataSheet.getRange(`A1:${lastCol}1`).format.font.color = HEADER_FONT;
dataSheet.getRange(`A1:${lastCol}1`).format.font.bold = true;
dataSheet.getRange(`A1:${lastCol}${rowCount + 1}`).format.borders = { preset: "all", style: "thin", color: BORDER };

headers.forEach((header, index) => {
  const col = columnLetter(index);
  const range = dataSheet.getRange(`${col}2:${col}${rowCount + 1}`);
  if (inputFields.has(header)) {
    inputCols.push(col);
    dataSheet.getRange(`${col}1`).format.fill.color = INPUT_FILL;
    dataSheet.getRange(`${col}1`).format.font.color = "#5B4300";
    range.format.fill.color = INPUT_FILL_SOFT;
  } else if (outputFields.has(header)) {
    dataSheet.getRange(`${col}1`).format.fill.color = OUTPUT_FILL;
    dataSheet.getRange(`${col}1`).format.font.color = "#1F4E78";
    range.format.fill.color = OUTPUT_FILL_SOFT;
  }
});

for (let row = 2; row <= rowCount + 1; row += 1) {
  dataSheet.getRange(`I${row}`).formulas = [[`=IF(OR($G${row}="",$F${row}=""),"",ROUND($G${row}*1000000/$F${row},2))`]];
  dataSheet.getRange(`J${row}`).formulas = [[`=IF(OR($H${row}="",$F${row}=""),"",ROUND($H${row}*1000000/$F${row},1))`]];
  dataSheet.getRange(`M${row}`).formulas = [[
    `=IF(OR(I${row}="",J${row}="",K${row}="",L${row}=""),"",ROUND(MIN(100,MAX(0,MIN(35,MAX(0,(I${row}-8)/6*35))+MIN(35,MAX(0,(J${row}-20)/16*35))+MIN(30,MAX(0,(((I${row}/K${row})-1)+((J${row}/L${row})-1))/0.25*30)))),0))`,
  ]];
  dataSheet.getRange(`N${row}`).formulas = [[`=IF($M${row}="","",IF($M${row}>=75,"hot",IF($M${row}>=55,"warn","good")))`]];
}

dataSheet.getRange(`E2:F${rowCount + 1}`).format.numberFormat = "0";
dataSheet.getRange(`G2:H${rowCount + 1}`).format.numberFormat = "0.00";
dataSheet.getRange(`I2:L${rowCount + 1}`).format.numberFormat = "0.0";
dataSheet.getRange(`M2:M${rowCount + 1}`).format.numberFormat = "0";
dataSheet.getRange(`O2:Z${rowCount + 1}`).format.numberFormat = "0";
dataSheet.getRange(`AA2:AE${rowCount + 1}`).format.numberFormat = "0";
dataSheet.getRange(`A1:${lastCol}${rowCount + 1}`).format.wrapText = false;
dataSheet.getRange(`A1:${lastCol}${rowCount + 1}`).format.autofitColumns();
dataSheet.getRange("A:A").format.columnWidthPx = 150;
dataSheet.getRange("B:B").format.columnWidthPx = 96;
dataSheet.getRange("C:D").format.columnWidthPx = 104;

dataSheet.getRange(`C2:C${rowCount + 1}`).dataValidation = { rule: { type: "list", values: ["Singapore", "Japan"] } };
dataSheet.getRange(`D2:D${rowCount + 1}`).dataValidation = { rule: { type: "list", values: ["Office", "Retail", "Logistics"] } };

dataSheet.getRange(`M2:M${rowCount + 1}`).conditionalFormats.add("colorScale", {
  criteria: [
    { type: "lowestValue", color: "#E8F5EF" },
    { type: "percentile", value: 55, color: "#FFF2D9" },
    { type: "highestValue", color: "#FAE6E2" },
  ],
});

guideSheet.getRange("A1:D1").values = [["Field", "Role", "Editable?", "Definition / Calculation"]];
guideSheet.getRange("A1:D1").format.fill.color = HEADER_FILL;
guideSheet.getRange("A1:D1").format.font.color = HEADER_FONT;
guideSheet.getRange("A1:D1").format.font.bold = true;
guideSheet.freezePanes.freezeRows(1);

const definitions = {
  name: "Full asset name.",
  label: "Short asset label used in charts and tiles.",
  country: "Country grouping.",
  type: "Asset type. Office includes business park assets in the dashboard.",
  gfa: "Gross floor area.",
  nla: "Net lettable area.",
  opex_m: "Monthly or period operating cost, in millions.",
  energy_gwh: "Energy consumption, in GWh.",
  cost_sqm: "Calculated as opex_m x 1,000,000 / NLA.",
  kwh_sqm: "Calculated as energy_gwh x 1,000,000 / NLA.",
  prev_cost_sqm: "Previous FY cost intensity baseline.",
  prev_kwh_sqm: "Previous FY energy intensity baseline.",
  score: "Calculated pressure score from cost intensity, energy intensity and YoY pressure.",
  status: "Calculated traffic-light status from score: hot >=75, warn >=55, otherwise good.",
};

const guideRows = headers.map((header) => [
  header,
  inputFields.has(header) ? "Input" : "Calculated output",
  inputFields.has(header) ? "Yes" : "No",
  definitions[header] ?? (header.startsWith("trend_")
    ? "Monthly heat / pressure index input for the last 12-month sparkline."
    : header.startsWith("cost_")
      ? "Cost composition percentage input."
      : "Dashboard field."),
]);
guideSheet.getRange("A2").write(guideRows);
guideSheet.getRange(`A2:D${guideRows.length + 1}`).format.borders = { preset: "all", style: "thin", color: BORDER };
guideSheet.getRange(`B2:C${guideRows.length + 1}`).format.font.bold = true;
for (let row = 2; row <= guideRows.length + 1; row += 1) {
  const isInput = guideSheet.getRange(`B${row}`).values[0][0] === "Input";
  guideSheet.getRange(`A${row}:D${row}`).format.fill.color = isInput ? INPUT_FILL_SOFT : OUTPUT_FILL_SOFT;
}
guideSheet.getRange("F1").values = [["Legend"]];
guideSheet.getRange("F1").format.font.bold = true;
guideSheet.getRange("F2:G3").values = [
  ["Editable input", "Pale yellow cells can be changed."],
  ["Calculated output", "Pale blue cells are formula-driven and locked."],
];
guideSheet.getRange("F2:F2").format.fill.color = INPUT_FILL;
guideSheet.getRange("F3:F3").format.fill.color = OUTPUT_FILL;
guideSheet.getRange("A:D").format.autofitColumns();
guideSheet.getRange("D:D").format.columnWidthPx = 620;
guideSheet.getRange("G:G").format.columnWidthPx = 360;
guideSheet.getRange("F:F").format.autofitColumns();

await fs.rm(tempDir, { recursive: true, force: true });
await fs.mkdir(tempDir, { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(tempXlsx);
await protectWorkbook("Portfolio Data", inputCols, 2, rowCount + 1);

console.log(`Saved ${outPath}`);
