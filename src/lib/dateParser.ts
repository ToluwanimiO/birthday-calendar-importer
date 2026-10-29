import type { CellValue } from "../types";

export type ParsedBirthday = {
  month: number;
  day: number;
  birthYear?: number;
};

export type DateParseResult =
  | { ok: true; value: ParsedBirthday; warning?: string }
  | { ok: false; error: string };

function isLeapYear(year: number): boolean {
  return year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0);
}

function validMonthDay(month: number, day: number, year = 2000): boolean {
  if (!Number.isInteger(month) || !Number.isInteger(day) || month < 1 || month > 12 || day < 1) {
    return false;
  }
  const days = [31, isLeapYear(year) ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  return day <= days[month - 1];
}

function fromDate(date: Date): DateParseResult {
  if (Number.isNaN(date.getTime())) {
    return { ok: false, error: "This date could not be read." };
  }
  const year = date.getFullYear();
  const month = date.getMonth() + 1;
  const day = date.getDate();
  if (!validMonthDay(month, day, year)) {
    return { ok: false, error: "This is not a valid calendar date." };
  }
  return { ok: true, value: { month, day, birthYear: year } };
}

/**
 * Parses spreadsheet values. Ambiguous numeric text is rejected unless the
 * user chooses a date order (DMY or MDY). ISO yyyy-mm-dd is unambiguous.
 */
export function parseBirthday(
  input: CellValue,
  dateOrder: "DMY" | "MDY" = "DMY",
): DateParseResult {
  if (input === null || input === undefined || input === "") {
    return { ok: false, error: "Birthday is blank." };
  }

  if (input instanceof Date) return fromDate(input);

  if (typeof input === "number" && Number.isFinite(input)) {
    // Excel's common 1900 date system. SheetJS normally returns raw serials
    // unless cellDates is enabled. Convert serial to UTC to avoid timezone drift.
    if (input < 1 || input > 2958465) {
      return { ok: false, error: "This number does not look like an Excel date." };
    }
    const date = new Date(Date.UTC(1899, 11, 30) + Math.floor(input) * 86400000);
    return fromDate(date);
  }

  const text = String(input).trim();
  if (!text) return { ok: false, error: "Birthday is blank." };

  // ISO format yyyy-mm-dd (or yyyy/m/d).
  const iso = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if (iso) {
    const year = Number(iso[1]);
    const month = Number(iso[2]);
    const day = Number(iso[3]);
    if (!validMonthDay(month, day, year)) return { ok: false, error: "This is not a valid date." };
    return { ok: true, value: { month, day, birthYear: year } };
  }

  // Named month dates with or without a year: "15 March 1998", "March 15", "Mar 15, 1998".
  const months: Record<string, number> = {
    jan: 1, january: 1, feb: 2, february: 2, mar: 3, march: 3, apr: 4, april: 4,
    may: 5, jun: 6, june: 6, jul: 7, july: 7, aug: 8, august: 8, sep: 9,
    sept: 9, september: 9, oct: 10, october: 10, nov: 11, november: 11,
    dec: 12, december: 12,
  };
  const namedText = text.toLowerCase().replace(/,/g, "").trim();
  const named = namedText.match(/^([a-z]+)\s+(\d{1,2})(?:\s+(\d{4}))?$|^(\d{1,2})\s+([a-z]+)(?:\s+(\d{4}))?$/);
  if (named) {
    const monthName = named[1] ?? named[5];
    const day = Number(named[2] ?? named[4]);
    const yearText = named[3] ?? named[6];
    const year = yearText ? Number(yearText) : undefined;
    const month = months[monthName.replace(/\.$/, "")] ?? months[monthName.slice(0, 3)];
    if (!month || !validMonthDay(month, day, year ?? 2000)) {
      return { ok: false, error: "Check the month and day." };
    }
    return { ok: true, value: { month, day, ...(year ? { birthYear: year } : {}) } };
  }

  // Numeric dates: d/m/yyyy or m/d/yyyy. Require a year for this format.
  const numeric = text.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/);
  if (numeric) {
    const first = Number(numeric[1]);
    const second = Number(numeric[2]);
    const year = Number(numeric[3]);
    const month = dateOrder === "DMY" ? second : first;
    const day = dateOrder === "DMY" ? first : second;
    if (!validMonthDay(month, day, year)) {
      return { ok: false, error: "This date is invalid under the selected date format." };
    }
    const ambiguous = first <= 12 && second <= 12 && first !== second;
    return {
      ok: true,
      value: { month, day, birthYear: year },
      warning: ambiguous ? `Interpreted as ${dateOrder === "DMY" ? "day/month/year" : "month/day/year"}.` : undefined,
    };
  }

  return {
    ok: false,
    error: "Could not recognize this date. Try YYYY-MM-DD, 15 March, or choose the correct numeric date format.",
  };
}

export function formatBirthday(record: { month: number | null; day: number | null; birthYear?: number }): string {
  if (!record.month || !record.day) return "";
  const date = new Date(Date.UTC(record.birthYear ?? 2000, record.month - 1, record.day));
  return new Intl.DateTimeFormat("en", {
    month: "long",
    day: "numeric",
    ...(record.birthYear ? { year: "numeric" } : {}),
    timeZone: "UTC",
  }).format(date);
}
