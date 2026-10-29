import ical, { ICalEventRepeatingFreq, ICalEventTransparency } from "ical-generator";
import type { BirthdayRecord } from "../types";

export type IcsOptions = {
  calendarName: string;
  includeBirthYear: boolean;
  eventPrefix: string;
};

function stableUid(record: BirthdayRecord): string {
  // Stable for this source row and birthday, to reduce accidental duplicate UIDs
  // when the same spreadsheet is exported repeatedly. This does not update events
  // already imported into a calendar; ICS import behavior is app-specific.
  const safeName = record.name.toLowerCase().trim().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
  return `birthday-${record.sourceRow}-${record.month}-${record.day}-${safeName || "person"}@birthday-calendar-importer.local`;
}

export function generateIcs(records: BirthdayRecord[], options: IcsOptions): string {
  const calendar = ical({
    name: options.calendarName.trim() || "Birthdays",
    prodId: { company: "Birthday Calendar Importer", product: "Birthday Calendar Importer", language: "EN" },
    timezone: "Etc/UTC",
  });

  for (const record of records) {
    if (!record.included || record.status === "error" || !record.month || !record.day) continue;

    // Use a leap year to validate Feb 29. The recurrence is yearly and the event
    // remains on Feb 29, occurring only in leap years in many calendar apps.
    const year = record.birthYear ?? 2000;
    const startDate = new Date(Date.UTC(year, record.month - 1, record.day, 0, 0, 0));
    const summaryName = options.includeBirthYear && record.birthYear
      ? `${record.name} (${record.birthYear})`
      : record.name;
    const summary = `${options.eventPrefix.trim()}${summaryName}`.trim();

    calendar.createEvent({
      id: stableUid(record),
      start: startDate,
      allDay: true,
      summary,
      description: record.birthYear
        ? `Birthday reminder${options.includeBirthYear ? ` · Born ${record.birthYear}` : ""}`
        : "Birthday reminder",
      repeating: { freq: ICalEventRepeatingFreq.YEARLY },
      transparency: ICalEventTransparency.TRANSPARENT,
    });
  }

  return calendar.toString();
}

export function downloadIcs(contents: string, fileName: string): void {
  const blob = new Blob([contents], { type: "text/calendar;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName.toLowerCase().endsWith(".ics") ? fileName : `${fileName}.ics`;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}
