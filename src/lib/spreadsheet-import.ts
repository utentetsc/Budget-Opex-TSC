import * as XLSX from "xlsx";
import type { Cinema, HistoricalCategory } from "./budget-model";

export interface WorkbookSheet {
  name: string;
  headers: string[];
  headerRow: number;
  rows: Record<string, unknown>[];
}

export interface WorkbookData {
  fileName: string;
  sheets: WorkbookSheet[];
}

export interface ImportMapping {
  cinemaField: string;
  categoryField: string;
  amountField: string;
  yearField: string;
}

export interface ImportSummary {
  cinemas: Cinema[];
  categories: HistoricalCategory[];
  years: number[];
  selectedYear: number;
  totalSpend: number;
  transactionCount: number;
  skippedAmountRows: number;
  excludedCenters: string[];
}

const HEADER_SCAN_LIMIT = 40;
const ODEON_NAME = "TSC Odeon Milano";

function textValue(value: unknown): string {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

function uniqueHeaders(values: unknown[]): string[] {
  const seen = new Map<string, number>();
  return values.map((value, index) => {
    const base = textValue(value) || `Colonna ${index + 1}`;
    const count = seen.get(base) ?? 0;
    seen.set(base, count + 1);
    return count === 0 ? base : `${base} (${count + 1})`;
  });
}

function likelyHeaderRow(row: unknown[]): boolean {
  const normalized = row.map((value) => textValue(value).toLowerCase()).join(" | ");
  return /commessa|cinema|centro/.test(normalized)
    && /importo|spesa|costo|budget/.test(normalized)
    && /anno|year/.test(normalized);
}

function findHeaderRow(rows: unknown[][]): number {
  const candidate = rows.findIndex((row, index) => index < HEADER_SCAN_LIMIT && likelyHeaderRow(row));
  return candidate >= 0 ? candidate : 0;
}

export async function readWorkbook(file: File): Promise<WorkbookData> {
  const bytes = await file.arrayBuffer();
  const workbook = XLSX.read(bytes, { type: "array", cellDates: true, raw: false });
  const sheets = workbook.SheetNames.map((name) => {
    const worksheet = workbook.Sheets[name];
    if (!worksheet) return { name, headers: [], headerRow: 0, rows: [] };
    const matrix = XLSX.utils.sheet_to_json<unknown[]>(worksheet, {
      header: 1,
      defval: null,
      raw: false,
      blankrows: false,
    });
    const headerRow = findHeaderRow(matrix);
    const headers = uniqueHeaders(matrix[headerRow] ?? []);
    const rows = matrix.slice(headerRow + 1)
      .filter((row) => row.some((cell) => cell !== null && cell !== undefined && textValue(cell) !== ""))
      .map((row) => Object.fromEntries(headers.map((header, index) => [header, row[index] ?? null])));
    return { name, headers, headerRow, rows };
  });

  return { fileName: file.name, sheets };
}

export function detectImportMapping(headers: string[]): Partial<ImportMapping> {
  const match = (pattern: RegExp) => headers.find((header) => pattern.test(header.toLowerCase()));
  const amountField = match(/importo.*fatturato/)
    ?? match(/importo.*verificato/)
    ?? match(/importo.*ordinato/)
    ?? match(/amount|spesa|costo/);
  return {
    cinemaField: match(/commessa.*sap|cinema|centro.*costo|sede/),
    categoryField: match(/causale/)
      ?? match(/categoria|category/)
      ?? match(/voce.*budget/)
      ?? match(/area.*riferimento/),
    amountField,
    yearField: match(/anno.*competenza|anno|year/),
  };
}

function parseYear(value: unknown): number | null {
  const raw = textValue(value);
  const candidate = raw.match(/(?:19|20)\d{2}/)?.[0] ?? raw;
  const year = Number(candidate);
  return Number.isInteger(year) && year >= 1900 && year <= 2200 ? year : null;
}

export function parseEuroAmount(value: unknown): number | null {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  let raw = textValue(value).replace(/[€\s\u00a0]/g, "");
  if (!raw) return null;
  const negative = raw.startsWith("(") && raw.endsWith(")");
  raw = raw.replace(/[()]/g, "").replace(/[^\d,.-]/g, "");
  if (!raw || raw === "-" || raw === "." || raw === ",") return null;

  const comma = raw.lastIndexOf(",");
  const dot = raw.lastIndexOf(".");
  if (comma >= 0 && dot >= 0) {
    const decimalSeparator = comma > dot ? "," : ".";
    const thousandsSeparator = decimalSeparator === "," ? "." : ",";
    raw = raw.split(thousandsSeparator).join("");
    if (decimalSeparator === ",") raw = raw.replace(",", ".");
  } else if (comma >= 0 || dot >= 0) {
    const separator = comma >= 0 ? "," : ".";
    const parts = raw.split(separator);
    const lastPart = parts.at(-1) ?? "";
    if (lastPart.length <= 2) {
      raw = `${parts.slice(0, -1).join("")}.${lastPart}`;
    } else {
      raw = parts.join("");
    }
  }

  const amount = Number(raw);
  if (!Number.isFinite(amount)) return null;
  return negative ? -amount : amount;
}

function makeId(value: string, suffix = ""): string {
  const slug = value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-|-$/g, "");
  return `${slug || "cinema"}${suffix}`;
}

function normalizedName(value: string): string {
  return value.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

function isHeadOffice(value: string): boolean {
  return /\b(sede|head office|hq|ufficio)\b/i.test(value);
}

export function summarizeImport(
  sheet: WorkbookSheet,
  mapping: ImportMapping,
  selectedYear: number,
  categoryGroups: Record<string, string> = {},
): ImportSummary {
  const years = [...new Set(sheet.rows
    .map((row) => parseYear(row[mapping.yearField]))
    .filter((year): year is number => year !== null))]
    .sort((a, b) => a - b);

  const centerNames = [...new Set(sheet.rows
    .map((row) => textValue(row[mapping.cinemaField]))
    .filter(Boolean))];
  const includedCenters = centerNames.filter((name) => !isHeadOffice(name));
  const excludedCenters = centerNames.filter(isHeadOffice);
  const cinemas = includedCenters.map((name, index) => ({
    id: makeId(name, `-${index + 1}`),
    name,
    attendance: null,
    areaSqm: null,
    openingYear: null,
    renovated: null,
    locationType: null,
    roofStatus: null,
    maintenanceResponsibility: null,
    factorScores: {},
  }));

  const hasOdeon = cinemas.some((cinema) => normalizedName(cinema.name).includes("odeonmilano"));
  if (!hasOdeon) {
    cinemas.push({
      id: "tsc-odeon-milano",
      name: ODEON_NAME,
      attendance: null,
      areaSqm: null,
      openingYear: null,
      renovated: null,
      locationType: null,
      roofStatus: null,
      maintenanceResponsibility: null,
      factorScores: {},
    });
  }

  const spendByCategory = new Map<string, number>();
  let totalSpend = 0;
  let transactionCount = 0;
  let skippedAmountRows = 0;

  for (const row of sheet.rows) {
    if (parseYear(row[mapping.yearField]) !== selectedYear) continue;
    const center = textValue(row[mapping.cinemaField]);
    if (!center || isHeadOffice(center)) continue;

    const amount = parseEuroAmount(row[mapping.amountField]);
    if (amount === null) {
      skippedAmountRows += 1;
      continue;
    }
    const categorySource = textValue(row[mapping.categoryField]) || "Non classificato";
    const category = textValue(categoryGroups[categorySource]) || categorySource;
    spendByCategory.set(category, (spendByCategory.get(category) ?? 0) + amount);
    totalSpend += amount;
    transactionCount += 1;
  }

  const categories = [...spendByCategory.entries()]
    .filter(([, spend]) => spend > 0)
    .map(([name, previousYearSpend], index) => ({
      id: makeId(name, `-${index + 1}`),
      name,
      previousYearSpend: Math.round(previousYearSpend * 100) / 100,
    }))
    .sort((a, b) => b.previousYearSpend - a.previousYearSpend);
  totalSpend = categories.reduce((sum, category) => sum + category.previousYearSpend, 0);

  return {
    cinemas,
    categories,
    years,
    selectedYear,
    totalSpend: Math.round(totalSpend * 100) / 100,
    transactionCount,
    skippedAmountRows,
    excludedCenters,
  };
}
