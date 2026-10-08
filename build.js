// Reads the local resident-list .xlsx and writes data.json with aggregated,
// anonymous KPIs only (no names, no birthdates) so that file is safe to commit
// and publish via GitHub Pages. Run this manually after updating the sheet:
//
//   npm run build
//
import { readdirSync, readFileSync, writeFileSync } from "node:fs";
import XLSX from "xlsx";

const SOURCE_PATTERN = /Bewonerslijst.*\.xlsx$/i;

function findSourceFile() {
  const candidates = readdirSync(".").filter(
    (f) => SOURCE_PATTERN.test(f) && !f.startsWith("~$"),
  );
  if (candidates.length === 0) {
    throw new Error("No *Bewonerslijst*.xlsx file found in this directory.");
  }
  // If there are several (e.g. old exports kept around), take the most recent.
  candidates.sort();
  return candidates[candidates.length - 1];
}

function ageOnDate(birthDate, onDate) {
  let age = onDate.getFullYear() - birthDate.getFullYear();
  const monthDay = (d) => d.getMonth() * 100 + d.getDate();
  if (monthDay(birthDate) > monthDay(onDate)) age -= 1;
  return age;
}

function median(values) {
  const sorted = [...values].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 0
    ? (sorted[mid - 1] + sorted[mid]) / 2
    : sorted[mid];
}

function round(n, decimals) {
  const f = 10 ** decimals;
  return Math.round(n * f) / f;
}

const sourceFile = findSourceFile();
const workbook = XLSX.read(readFileSync(sourceFile), { cellDates: true });
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const rows = XLSX.utils.sheet_to_json(sheet, { raw: false, cellDates: true });

const years = Object.keys(rows[0])
  .filter((k) => /^\d{4}$/.test(k))
  .map(Number)
  .sort((a, b) => a - b);

const residentsByYear = {};
for (const year of years) {
  residentsByYear[year] = rows.filter(
    (r) => String(r[String(year)]).trim().toLowerCase() === "x",
  );
}

const series = {
  avgAge: [],
  medianAge: [],
  residents: [],
  unitsOccupied: [],
  avgHouseholdSize: [],
  moveIns: [],
  moveOuts: [],
};

for (let i = 0; i < years.length; i++) {
  const year = years[i];
  const present = residentsByYear[year];
  const refDate = new Date(year, 11, 31); // age as of Dec 31 of that year

  const ages = present.map((r) => ageOnDate(new Date(r.Geboortedatum), refDate));
  const units = new Set(present.map((r) => r.Unit));

  series.avgAge.push(round(ages.reduce((a, b) => a + b, 0) / ages.length, 1));
  series.medianAge.push(round(median(ages), 1));
  series.residents.push(present.length);
  series.unitsOccupied.push(units.size);
  series.avgHouseholdSize.push(round(present.length / units.size, 2));

  if (i === 0) {
    series.moveIns.push(null);
    series.moveOuts.push(null);
  } else {
    const prevNames = new Set(residentsByYear[years[i - 1]].map((r) => r.Naam + "|" + r.Geboortedatum));
    const currNames = new Set(present.map((r) => r.Naam + "|" + r.Geboortedatum));
    const movedIn = [...currNames].filter((k) => !prevNames.has(k)).length;
    const movedOut = [...prevNames].filter((k) => !currNames.has(k)).length;
    series.moveIns.push(movedIn);
    series.moveOuts.push(movedOut);
  }
}

const data = {
  generatedAt: new Date().toISOString(),
  ageReference: "Age computed as of December 31 of each year.",
  years,
  series,
};

writeFileSync("data.json", JSON.stringify(data, null, 2));
console.log(`Wrote data.json from ${sourceFile} (${years.length} years, ${rows.length} residents total).`);
