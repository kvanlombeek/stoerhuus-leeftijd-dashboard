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

// Reads a "Geslacht" column (M/V, Man/Vrouw, ...). Returns "M", "V", or null
// (missing/unrecognized) — null rows are counted but left out of the pyramid.
function normalizeGender(raw) {
  const first = String(raw ?? "").trim().toLowerCase().charAt(0);
  if (first === "m") return "M";
  if (first === "v") return "V";
  return null;
}

function round(n, decimals) {
  const f = 10 ** decimals;
  return Math.round(n * f) / f;
}

// Linear-interpolation quantile (R type 7 / numpy default) on a sorted array.
function quantile(sorted, q) {
  const pos = (sorted.length - 1) * q;
  const base = Math.floor(pos);
  const rest = pos - base;
  return sorted[base + 1] !== undefined
    ? sorted[base] + rest * (sorted[base + 1] - sorted[base])
    : sorted[base];
}

function boxplotStats(ages) {
  if (ages.length === 0) return null;
  const sorted = [...ages].sort((a, b) => a - b);
  const q1 = quantile(sorted, 0.25);
  const q3 = quantile(sorted, 0.75);
  const iqr = q3 - q1;
  const lowFence = q1 - 1.5 * iqr;
  const highFence = q3 + 1.5 * iqr;
  const inRange = sorted.filter((a) => a >= lowFence && a <= highFence);
  const outliers = sorted.filter((a) => a < lowFence || a > highFence);
  return {
    min: sorted[0],
    q1: round(q1, 1),
    median: round(quantile(sorted, 0.5), 1),
    q3: round(q3, 1),
    max: sorted[sorted.length - 1],
    whiskerLow: inRange[0],
    whiskerHigh: inRange[inRange.length - 1],
    outliers,
  };
}

const sourceFile = findSourceFile();
const workbook = XLSX.read(readFileSync(sourceFile), { cellDates: true });
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const rows = XLSX.utils.sheet_to_json(sheet, { cellDates: true });

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

const totalUnits = new Set(rows.map((r) => r.Unit)).size;

const genderKey = Object.keys(rows[0]).find((k) => k.trim().toLowerCase() === "geslacht");
if (!genderKey) {
  console.warn('No "Geslacht" column found — the gender-split age pyramid will be empty until one is added.');
}

// Age-histogram bucket edges: decades, sized to the oldest age seen anywhere.
const allAgesEver = years.flatMap((year) => {
  const refDate = new Date(year, 11, 31);
  return residentsByYear[year].map((r) => ageOnDate(new Date(r.Geboortedatum), refDate));
});
const maxBucketStart = Math.floor(Math.max(...allAgesEver) / 10) * 10;
const bucketEdges = [];
for (let start = 0; start <= maxBucketStart; start += 10) bucketEdges.push(start);
const bucketLabels = bucketEdges.map((start, i) =>
  i === bucketEdges.length - 1 ? `${start}+` : `${start}-${start + 9}`,
);

function bucketIndex(age) {
  return Math.min(Math.floor(age / 10), bucketLabels.length - 1);
}

function histogramFor(ages) {
  const counts = new Array(bucketLabels.length).fill(0);
  for (const age of ages) counts[bucketIndex(age)] += 1;
  return counts;
}

// Per-gender bucket counts, for the population-pyramid chart.
function histogramByGenderFor(present, ages) {
  const male = new Array(bucketLabels.length).fill(0);
  const female = new Array(bucketLabels.length).fill(0);
  let unknown = 0;
  present.forEach((r, idx) => {
    const gender = genderKey ? normalizeGender(r[genderKey]) : null;
    if (gender === "M") male[bucketIndex(ages[idx])] += 1;
    else if (gender === "V") female[bucketIndex(ages[idx])] += 1;
    else unknown += 1;
  });
  return { male, female, unknown };
}

const series = {
  avgAge: [],
  medianAge: [],
  residents: [],
  unitsOccupied: [],
  avgHouseholdSize: [],
  children: [],
  moveIns: [],
  moveOuts: [],
  maleCount: [],
  femaleCount: [],
  sexRatio: [],
};
const ageHistogramByYear = [];
const ageBoxplotByYear = [];
const agePyramidByYear = [];
let unknownGenderTotal = 0;

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
  series.children.push(ages.filter((age) => age < 18).length);
  ageHistogramByYear.push(histogramFor(ages));

  const maleAges = [];
  const femaleAges = [];
  present.forEach((r, idx) => {
    const gender = genderKey ? normalizeGender(r[genderKey]) : null;
    if (gender === "M") maleAges.push(ages[idx]);
    else if (gender === "V") femaleAges.push(ages[idx]);
  });
  ageBoxplotByYear.push({ male: boxplotStats(maleAges), female: boxplotStats(femaleAges) });
  series.maleCount.push(maleAges.length);
  series.femaleCount.push(femaleAges.length);
  series.sexRatio.push(femaleAges.length > 0 ? round((maleAges.length / femaleAges.length) * 100, 0) : null);

  const pyramid = histogramByGenderFor(present, ages);
  agePyramidByYear.push(pyramid);
  unknownGenderTotal += pyramid.unknown;

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
  totalUnits,
  years,
  series,
  ageHistogram: {
    buckets: bucketLabels,
    countsByYear: ageHistogramByYear,
  },
  ageBoxplot: ageBoxplotByYear,
  agePyramid: {
    buckets: bucketLabels,
    byYear: agePyramidByYear,
  },
};

writeFileSync("data.json", JSON.stringify(data, null, 2));
console.log(`Wrote data.json from ${sourceFile} (${years.length} years, ${rows.length} residents total).`);
if (unknownGenderTotal > 0) {
  console.warn(`${unknownGenderTotal} resident-year entries have no recognized gender (expected "M" or "V") and are left out of the age pyramid.`);
}
