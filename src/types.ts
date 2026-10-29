export type CellValue = string | number | boolean | Date | null | undefined;

export type BirthdayRecord = {
  id: string;
  name: string;
  month: number | null;
  day: number | null;
  birthYear?: number;
  included: boolean;
  status: "valid" | "warning" | "error";
  issues: string[];
  sourceRow: number;
  originalBirthday: string;
};

export type ParsedWorkbook = {
  fileName: string;
  sheetName: string;
  headers: string[];
  rows: CellValue[][];
};
