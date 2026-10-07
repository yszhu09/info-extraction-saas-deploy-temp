import * as XLSX from "@e965/xlsx";
import type { SourceTable } from "@/lib/types";

const MAX_SHEETS = 3;

export const SPREADSHEET_ACCEPT = [
  ".xlsx",
  ".xls",
  ".csv",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "application/vnd.ms-excel",
  "text/csv",
].join(",");

export function isSpreadsheetFile(file: File) {
  const lowerName = file.name.toLowerCase();
  return (
    lowerName.endsWith(".xlsx") ||
    lowerName.endsWith(".xls") ||
    lowerName.endsWith(".csv") ||
    file.type === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" ||
    file.type === "application/vnd.ms-excel" ||
    file.type === "text/csv"
  );
}

function formatCellValue(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === "object") {
    const record = value as Record<string, unknown>;
    if (typeof record.text === "string") return record.text;
    if (typeof record.result === "string") return record.result;
    if (
      typeof record.richText !== "undefined" &&
      Array.isArray(record.richText)
    ) {
      return record.richText
        .map((part) => String((part as Record<string, unknown>).text ?? ""))
        .join("");
    }
    return JSON.stringify(value);
  }
  return String(value).trim();
}

export async function parseSpreadsheetFile(file: File): Promise<SourceTable[]> {
  const workbook = XLSX.read(await file.arrayBuffer(), {
    type: "array",
    cellDates: true,
    raw: false,
  });

  return workbook.SheetNames.slice(0, MAX_SHEETS).map((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils
      .sheet_to_json<unknown[]>(worksheet, {
        header: 1,
        blankrows: false,
        defval: "",
        raw: false,
      })
      .map((row) => row.map(formatCellValue))
      .filter((row) => row.some(Boolean));

    const [firstRow = []] = rows;
    const columns = firstRow.map(
      (value, index) => value || `Column ${index + 1}`,
    );

    return {
      sheetName,
      columns: columns.length ? columns : ["字段", "值"],
      rows: rows.slice(1),
    };
  });
}
