# Birthday Calendar Importer

A browser-based React + TypeScript app that reads Excel/CSV birthday lists, lets users review the rows, and downloads an iCalendar (`.ics`) file.

## Requirements

- Node.js 20 or newer recommended
- npm

## Run locally

```bash
npm install
npm run dev
```

Open the local URL printed by Vite.

## Production build

```bash
npm run build
npm run preview
```

The generated static site is in `dist/`, suitable for Netlify or another static host.

## Input format

The app supports `.xlsx`, `.xls`, and `.csv` files. It reads the first row as column headers. Select the worksheet and map the name and birthday columns.

Recognized date examples include:
- Excel date cells
- `1998-03-15`
- `15 March 1998`
- `February 29`
- `15/03/1998` with the DMY setting

Numeric text dates are interpreted using the chosen date order. Dates such as `03/04/1998` are ambiguous across locales, so verify the date order before previewing.

## Calendar export

The app creates yearly recurring all-day birthday events in an `.ics` file. To import, use the import function in your calendar application. Exact import steps and recurring-event handling vary by application. Importing the same file repeatedly can create duplicates.

## Privacy

The workbook is read locally in the browser. The app does not upload it to a backend. If deployed, the optional Google Fonts import in `src/styles.css` makes requests to Google Fonts; remove that `@import` if you want to avoid that third-party request too.

## Known MVP limitations

- Date formats that are not recognized need to be corrected in the preview.
- For numeric dates like `03/04/1998`, choose DMY or MDY in the mapping section.
- February 29 remains February 29 for yearly recurrence. Some calendar apps may handle non-leap years differently.
- This app exports events; it does not sync with or update events already imported into a calendar.
- Stable UIDs help identify events in iCalendar data, but some calendar importers may still create duplicates on repeated imports.
