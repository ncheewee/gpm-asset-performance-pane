import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const csvPath = new URL("../data/portfolio-data.csv", import.meta.url);
const outPath = new URL("../data/portfolio-data.xlsx", import.meta.url);

const csvText = await fs.readFile(csvPath, "utf8");
const workbook = await Workbook.fromCSV(csvText, { sheetName: "Portfolio Data" });

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outPath);

console.log(`Saved ${outPath.pathname}`);
