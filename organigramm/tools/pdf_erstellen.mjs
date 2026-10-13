// Erzeugt das PDF aus organigramm.html (optional mit eigener Excel-Datei).
//   node tools/pdf_erstellen.mjs [daten.xlsx] [A3|A4] [ausgabe.pdf]
// Voraussetzung: Node.js + Playwright (npm i playwright). Alternativ einfach im Browser „Drucken / PDF“.
import { chromium } from "playwright";
import { fileURLToPath, pathToFileURL } from "node:url";
import path from "node:path";

const here = path.dirname(fileURLToPath(import.meta.url));
const [xlsx, format = "A3", out = path.join(here, "..", "Organigramm.pdf")] = process.argv.slice(2);

const browser = await chromium.launch();
const page = await browser.newPage();
await page.goto(pathToFileURL(path.join(here, "..", "organigramm.html")).href);
if (xlsx) {
  await page.setInputFiles("#file", path.resolve(xlsx));
  await page.waitForFunction(() => !document.querySelector("#status").classList.contains("demo"));
}
await page.selectOption("#format", format);
await page.emulateMedia({ media: "print" });
await page.evaluate(() => window.fitForPrint());
await page.pdf({ path: out, preferCSSPageSize: true, printBackground: true });
await browser.close();
console.log("PDF geschrieben:", out);
