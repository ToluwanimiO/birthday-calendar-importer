import * as XLSX from "xlsx";
import type { CellValue, ParsedWorkbook } from "../types";

const ACCEPTED_EXTENSIONS = [".xlsx", ".xls", ".csv"];

export async function readWorkbook(file: File): Promise<{ sheets: ParsedWorkbook[] }> {
  const lowerName = file.name.toLowerCase();
  if (!ACCEPTED_EXTENSIONS.some((extension) => lowerName.endsWith(extension))) {
    throw new Error("Please choose an Excel (.xlsx, .xls) or CSV file.");
  }

  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, {
    type: "array",
    cellDates: true,
    raw: true,
  });

  const sheets = workbook.SheetNames.map((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const matrix = XLSX.utils.sheet_to_json<CellValue[]>(sheet, {
      header: 1,
      defval: "",
      raw: true,
      blankrows: false,
    });

    const firstNonEmptyIndex = matrix.findIndex((row) =>
      row.some((cell) => String(cell ?? "").trim() !== ""),
    );
    const headerRow = firstNonEmptyIndex >= 0 ? matrix[firstNonEmptyIndex] : [];
    const headers = headerRow.map((cell, index) => {
      const value = String(cell ?? "").trim();
      return value || `Column ${index + 1}`;
    });

    return {
      fileName: file.name,
      sheetName,
      headers,
      rows: matrix.slice(firstNonEmptyIndex + 1),
    };
  });

  return { sheets };
}
