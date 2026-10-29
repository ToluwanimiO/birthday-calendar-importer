import { useMemo, useRef, useState } from "react";
import type { ChangeEvent } from "react";
import { readWorkbook } from "./lib/excelParser";
import { formatBirthday, parseBirthday } from "./lib/dateParser";
import { downloadIcs, generateIcs } from "./lib/generateIcs";
import type { BirthdayRecord, CellValue, ParsedWorkbook } from "./types";

type DateOrder = "DMY" | "MDY";

function guessColumn(headers: string[], candidates: string[]): string {
  return headers.find((header) =>
    candidates.some((candidate) => header.toLowerCase().includes(candidate)),
  ) ?? "";
}

function cellToText(value: CellValue): string {
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  return String(value ?? "").trim();
}

function makeRecords(
  sheet: ParsedWorkbook,
  nameColumn: string,
  birthdayColumn: string,
  dateOrder: DateOrder,
): BirthdayRecord[] {
  const nameIndex = sheet.headers.indexOf(nameColumn);
  const birthdayIndex = sheet.headers.indexOf(birthdayColumn);

  return sheet.rows.map((row, index) => {
    const name = cellToText(row[nameIndex]);
    const originalBirthday = cellToText(row[birthdayIndex]);
    const parsed = parseBirthday(row[birthdayIndex], dateOrder);
    const issues: string[] = [];

    if (!name) issues.push("Name is blank.");
    if (!parsed.ok) issues.push(parsed.error);
    if (parsed.ok && parsed.warning) issues.push(parsed.warning);

    const status: BirthdayRecord["status"] = !name || !parsed.ok
      ? "error"
      : issues.length > 0
        ? "warning"
        : "valid";

    return {
      id: `${sheet.sheetName}-row-${index + 2}`,
      name,
      month: parsed.ok ? parsed.value.month : null,
      day: parsed.ok ? parsed.value.day : null,
      birthYear: parsed.ok ? parsed.value.birthYear : undefined,
      included: Boolean(name) && parsed.ok,
      status,
      issues,
      sourceRow: index + 2,
      originalBirthday,
    };
  }).filter((record) => record.name || record.originalBirthday);
}

export default function App() {
  const inputRef = useRef<HTMLInputElement>(null);
  const [fileName, setFileName] = useState("");
  const [sheets, setSheets] = useState<ParsedWorkbook[]>([]);
  const [sheetIndex, setSheetIndex] = useState(0);
  const [nameColumn, setNameColumn] = useState("");
  const [birthdayColumn, setBirthdayColumn] = useState("");
  const [dateOrder, setDateOrder] = useState<DateOrder>("DMY");
  const [records, setRecords] = useState<BirthdayRecord[]>([]);
  const [calendarName, setCalendarName] = useState("Birthdays");
  const [eventPrefix, setEventPrefix] = useState("");
  const [includeBirthYear, setIncludeBirthYear] = useState(false);
  const [error, setError] = useState("");
  const [notice, setNotice] = useState("");
  const [busy, setBusy] = useState(false);
  const [hasMapped, setHasMapped] = useState(false);

  const activeSheet = sheets[sheetIndex];

  const duplicateIds = useMemo(() => {
    const counts = new Map<string, number>();
    records.forEach((record) => {
      if (record.name && record.month && record.day) {
        const key = `${record.name.toLowerCase().trim()}|${record.month}|${record.day}`;
        counts.set(key, (counts.get(key) ?? 0) + 1);
      }
    });
    return new Set(records.filter((record) => {
      if (!record.name || !record.month || !record.day) return false;
      return (counts.get(`${record.name.toLowerCase().trim()}|${record.month}|${record.day}`) ?? 0) > 1;
    }).map((record) => record.id));
  }, [records]);

  const validCount = records.filter((record) => record.status !== "error" && record.month && record.day).length;
  const includedCount = records.filter((record) => record.included && record.status !== "error" && record.month && record.day).length;
  const errorCount = records.filter((record) => record.status === "error").length;

  async function handleFile(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    setError("");
    setNotice("");
    setBusy(true);
    try {
      const result = await readWorkbook(file);
      if (result.sheets.length === 0 || result.sheets.every((sheet) => sheet.headers.length === 0)) {
        throw new Error("This file does not contain a readable table.");
      }
      setFileName(file.name);
      setSheets(result.sheets);
      setSheetIndex(0);
      const first = result.sheets[0];
      const name = guessColumn(first.headers, ["name", "full name", "person"]);
      const birthday = guessColumn(first.headers, ["birthday", "birth date", "date of birth", "dob", "date"]);
      setNameColumn(name);
      setBirthdayColumn(birthday);
      setRecords([]);
      setHasMapped(false);
      setNotice(`Loaded ${file.name}. Choose the worksheet and columns, then preview the birthdays.`);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "Could not read this file.");
      setSheets([]);
      setRecords([]);
    } finally {
      setBusy(false);
      event.target.value = "";
    }
  }

  function changeSheet(value: string) {
    const index = Number(value);
    const next = sheets[index];
    setSheetIndex(index);
    setNameColumn(guessColumn(next.headers, ["name", "full name", "person"]));
    setBirthdayColumn(guessColumn(next.headers, ["birthday", "birth date", "date of birth", "dob", "date"]));
    setRecords([]);
    setHasMapped(false);
    setNotice("");
  }

  function previewRecords() {
    setError("");
    setNotice("");
    if (!activeSheet) {
      setError("Upload a spreadsheet first.");
      return;
    }
    if (!nameColumn || !birthdayColumn) {
      setError("Select both the name column and birthday column.");
      return;
    }
    if (nameColumn === birthdayColumn) {
      setError("Name and birthday must use different columns.");
      return;
    }
    const nextRecords = makeRecords(activeSheet, nameColumn, birthdayColumn, dateOrder);
    if (nextRecords.length === 0) {
      setError("No data rows were found under the selected columns.");
      return;
    }
    setRecords(nextRecords);
    setHasMapped(true);
    setNotice(`Preview ready: ${nextRecords.length} rows found. Review warnings and errors before exporting.`);
  }

  function updateRecord(id: string, updates: Partial<BirthdayRecord>) {
    setRecords((current) => current.map((record) => {
      if (record.id !== id) return record;
      const updated = { ...record, ...updates };
      if (updates.month !== undefined || updates.day !== undefined || updates.birthYear !== undefined || updates.name !== undefined) {
        const issues: string[] = [];
        if (!updated.name.trim()) issues.push("Name is blank.");
        const month = updated.month;
        const day = updated.day;
        if (!month || !day) issues.push("Choose a valid birthday.");
        else {
          const leap = month === 2 && day === 29;
          const testDate = new Date(Date.UTC(updated.birthYear ?? (leap ? 2000 : 2001), month - 1, day));
          if (testDate.getUTCMonth() + 1 !== month || testDate.getUTCDate() !== day) {
            issues.push("This is not a valid date.");
          }
        }
        updated.issues = issues;
        updated.status = issues.length ? "error" : "valid";
        if (issues.length) updated.included = false;
      }
      return updated;
    }));
  }

  function editDate(id: string, value: string) {
    const parsed = parseBirthday(value, dateOrder);
    if (!parsed.ok) {
      setRecords((current) => current.map((record) => record.id === id
        ? { ...record, originalBirthday: value, month: null, day: null, birthYear: undefined, status: "error", issues: [parsed.error], included: false }
        : record));
      return;
    }
    setRecords((current) => current.map((record) => record.id === id
      ? {
          ...record,
          originalBirthday: value,
          month: parsed.value.month,
          day: parsed.value.day,
          birthYear: parsed.value.birthYear,
          status: parsed.warning ? "warning" : "valid",
          issues: parsed.warning ? [parsed.warning] : [],
          included: Boolean(record.name.trim()),
        }
      : record));
  }

  function exportCalendar() {
    setError("");
    setNotice("");
    const eligible = records.filter((record) => record.included && record.status !== "error" && record.month && record.day);
    if (eligible.length === 0) {
      setError("There are no valid, included birthdays to export.");
      return;
    }
    try {
      const contents = generateIcs(eligible, { calendarName, includeBirthYear, eventPrefix });
      const safeName = (calendarName.trim() || "birthdays").replace(/[^a-z0-9-_]+/gi, "-").replace(/^-|-$/g, "");
      downloadIcs(contents, `${safeName || "birthdays"}.ics`);
      setNotice(`Downloaded calendar file with ${eligible.length} birthday event${eligible.length === 1 ? "" : "s"}. Import it into your calendar app.`);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "Could not generate the calendar file.");
    }
  }

  function loadSample() {
    const headers = ["Full Name", "Birthday", "Notes"];
    const sampleRows: CellValue[][] = [
      ["Ada Okafor", "15 March 1998", "Team"],
      ["Chidi Bello", "03/04/2001", "Ambiguous numeric date"],
      ["Maya James", "February 29", "Leap day"],
      ["", "12/05/1995", "Missing name"],
      ["Ada Okafor", "15 March 1998", "Possible duplicate"],
    ];
    const sampleSheet: ParsedWorkbook = {
      fileName: "sample-birthdays.xlsx",
      sheetName: "Birthdays",
      headers,
      rows: sampleRows,
    };
    setFileName(sampleSheet.fileName);
    setSheets([sampleSheet]);
    setSheetIndex(0);
    setNameColumn("Full Name");
    setBirthdayColumn("Birthday");
    setRecords([]);
    setHasMapped(false);
    setError("");
    setNotice("Sample data loaded. Preview it to see how the importer handles real-world cases.");
  }

  return (
    <main className="app-shell">
      {/* <header className="topbar">
        <a className="brand" href="/" aria-label="Birthday Calendar Importer home">
          <span className="brand-icon">✳</span>
          <span>birthday<span className="brand-light">book</span></span>
        </a>
        <span className="privacy-label"><span className="privacy-dot" /> Files stay on your device</span>
      </header> */}

      <section className="hero">
        <h1>Birthday Calendar Importer</h1>
        <p className="hero-copy">
          Turn a spreadsheet of birthdays into a calendar file you can import.
        </p>
      </section>

      {(error || notice) && (
        <div className={`message ${error ? "message-error" : "message-success"}`} role={error ? "alert" : "status"}>
          <span>{error ? "!" : "✓"}</span>
          <p>{error || notice}</p>
          <button className="message-close" onClick={() => { setError(""); setNotice(""); }} aria-label="Dismiss message">×</button>
        </div>
      )}

      <section className="workspace">
        <div className="step-heading">
          <span className="step-number">01</span>
          <div>
            <h2>Upload your spreadsheet</h2>
            <p>Start with an Excel file containing names and birthdays.</p>
          </div>
        </div>

        <div className="upload-panel">
          <input ref={inputRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleFile} hidden />
          <div className="upload-icon"><FileIcon /></div>
          <h3>{fileName || "Drop your birthday list here"}</h3>
          <p>Excel workbook (.xlsx, .xls) or CSV file</p>
          <button className="button button-primary" onClick={() => inputRef.current?.click()} disabled={busy}>
            {busy ? "Reading file…" : fileName ? "Choose another file" : "Choose a file"}
          </button>
          <button className="text-button" onClick={loadSample}>Try sample data instead</button>
        </div>

        {activeSheet && (
          <div className="mapping-panel">
            <div className="panel-title-row">
              <div>
                <h3>Match your columns</h3>
                <p>Tell us where each piece of information lives.</p>
              </div>
              <span className="file-pill">{activeSheet.headers.length} columns</span>
            </div>
            {sheets.length > 1 && (
              <label className="field">
                <span>Worksheet</span>
                <select value={sheetIndex} onChange={(event) => changeSheet(event.target.value)}>
                  {sheets.map((sheet, index) => <option key={sheet.sheetName} value={index}>{sheet.sheetName}</option>)}
                </select>
              </label>
            )}
            <div className="mapping-grid">
              <label className="field">
                <span>Name column</span>
                <select value={nameColumn} onChange={(event) => setNameColumn(event.target.value)}>
                  <option value="">Choose a column</option>
                  {activeSheet.headers.map((header) => <option key={header} value={header}>{header}</option>)}
                </select>
              </label>
              <label className="field">
                <span>Birthday column</span>
                <select value={birthdayColumn} onChange={(event) => setBirthdayColumn(event.target.value)}>
                  <option value="">Choose a column</option>
                  {activeSheet.headers.map((header) => <option key={header} value={header}>{header}</option>)}
                </select>
              </label>
            </div>
            <div className="format-row">
              <label className="field field-inline">
                <span>Numeric date format</span>
                <select value={dateOrder} onChange={(event) => setDateOrder(event.target.value as DateOrder)}>
                  <option value="DMY">DD/MM/YYYY (e.g. 03/04 = 3 April)</option>
                  <option value="MDY">MM/DD/YYYY (e.g. 03/04 = March 4)</option>
                </select>
              </label>
              <span className="hint">Used for dates like 03/04/1998</span>
            </div>
            <button className="button button-primary" onClick={previewRecords}>Preview birthdays <span aria-hidden="true">→</span></button>
          </div>
        )}

        {hasMapped && records.length > 0 && (
          <>
            <div className="step-heading step-heading-spaced">
              <span className="step-number">02</span>
              <div>
                <h2>Review your birthdays</h2>
                <p>Fix errors and uncheck anyone you don't want to include.</p>
              </div>
            </div>

            <div className="summary-grid">
              <div className="summary-card"><span>Rows found</span><strong>{records.length}</strong></div>
              <div className="summary-card"><span>Valid birthdays</span><strong>{validCount}</strong></div>
              <div className="summary-card"><span>Needs attention</span><strong className={errorCount ? "number-warning" : ""}>{errorCount}</strong></div>
            </div>

            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th className="include-col">Include</th>
                    <th>Name</th>
                    <th>Birthday</th>
                    <th>Status</th>
                  </tr>
                </thead>
                <tbody>
                  {records.map((record) => (
                    <tr key={record.id} className={record.status === "error" ? "row-error" : ""}>
                      <td>
                        <input
                          type="checkbox"
                          aria-label={`Include ${record.name || `row ${record.sourceRow}`}`}
                          checked={record.included}
                          disabled={record.status === "error"}
                          onChange={(event) => updateRecord(record.id, { included: event.target.checked })}
                        />
                      </td>
                      <td>
                        <input
                          className="table-input name-input"
                          aria-label={`Name in row ${record.sourceRow}`}
                          value={record.name}
                          onChange={(event) => updateRecord(record.id, { name: event.target.value })}
                        />
                        <small>Row {record.sourceRow}{duplicateIds.has(record.id) ? " · Possible duplicate" : ""}</small>
                      </td>
                      <td>
                        <input
                          className="table-input date-input"
                          aria-label={`Birthday in row ${record.sourceRow}`}
                          value={record.originalBirthday}
                          placeholder="e.g. 15 March 1998"
                          onChange={(event) => editDate(record.id, event.target.value)}
                          onBlur={(event) => {
                            if (event.target.value.trim()) editDate(record.id, event.target.value);
                          }}
                        />
                        {record.month && record.day && <small>{formatBirthday(record)}</small>}
                      </td>
                      <td>
                        <span className={`status-pill status-${record.status}`}>
                          <span />{record.status === "valid" ? "Ready" : record.status === "warning" ? "Check" : "Fix"}
                        </span>
                        {record.issues.length > 0 && <small className="issue-text">{record.issues.join(" ")}</small>}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            {duplicateIds.size > 0 && (
              <div className="duplicate-note"><span>↗</span> {duplicateIds.size} rows may be duplicates. Review them before exporting; they haven't been removed automatically.</div>
            )}

            <div className="step-heading step-heading-spaced">
              <span className="step-number">03</span>
              <div>
                <h2>Prepare your calendar</h2>
                <p>Choose how your events should appear.</p>
              </div>
            </div>

            <div className="export-panel">
              <div className="export-preview">
                <div className="calendar-tile"><span>MAR</span><strong>15</strong></div>
                <div>
                  <span className="preview-kicker">YOUR CALENDAR FILE</span>
                  <h3>{calendarName.trim() || "Birthdays"}</h3>
                  <p>{includedCount} recurring all-day birthday events</p>
                </div>
                <span className="ics-badge">.ICS</span>
              </div>
              <div className="export-fields">
                <label className="field">
                  <span>Calendar name</span>
                  <input value={calendarName} onChange={(event) => setCalendarName(event.target.value)} placeholder="e.g. Friends & Family" />
                </label>
                <label className="field">
                  <span>Optional event prefix</span>
                  <input value={eventPrefix} onChange={(event) => setEventPrefix(event.target.value)} placeholder="e.g. 🎂 " />
                  <small>Added before each person's name.</small>
                </label>
                <label className="toggle-row">
                  <span>
                    <strong>Include birth year in event title</strong>
                    <small>For example, “Ada Okafor (1998)”</small>
                  </span>
                  <input type="checkbox" checked={includeBirthYear} onChange={(event) => setIncludeBirthYear(event.target.checked)} />
                </label>
              </div>
              <button className="button button-primary button-wide" onClick={exportCalendar} disabled={includedCount === 0}>
                <DownloadIcon /> Download {includedCount} birthdays <span aria-hidden="true">→</span>
              </button>
              <p className="export-footnote"><LockIcon /> Your spreadsheet is processed in this browser. The downloaded file contains the birthdays you selected.</p>
            </div>

            <div className="howto">
              <h3>What happens next?</h3>
              <div className="howto-grid">
                <div><span>1</span><p>Download your <strong>.ics</strong> file.</p></div>
                <div><span>2</span><p>Open your calendar app and find its import option.</p></div>
                <div><span>3</span><p>Import the file and check a few birthdays.</p></div>
              </div>
              <p className="fine-print">Calendar apps differ in how they import recurring events. Importing the same file more than once may create duplicates.</p>
            </div>
          </>
        )}
      </section>
{/* 
      <footer>
        <span>birthday<span className="brand-light">book</span></span>
        <span>Made to make remembering easier.</span>
      </footer> */}
    </main>
  );
}

function FileIcon() {
  return <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M13.5 2.75H6.75A2.25 2.25 0 0 0 4.5 5v14a2.25 2.25 0 0 0 2.25 2.25h10.5A2.25 2.25 0 0 0 19.5 19V8.75Z"/><path d="M13.5 2.75V8.75H19.5M8 12h8M8 15.5h8M8 8.5h2"/></svg>;
}
function DownloadIcon() {
  return <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3.5v11m0 0 4-4m-4 4-4-4M4.5 17v2.25A1.75 1.75 0 0 0 6.25 21h11.5a1.75 1.75 0 0 0 1.75-1.75V17"/></svg>;
}
function LockIcon() {
  return <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="5" y="10" width="14" height="11" rx="2"/><path d="M8 10V7a4 4 0 1 1 8 0v3m-4 4v3"/></svg>;
}
