# CSV to Excel Converter — specification

**Tagline:** Turn CSV files into a real XLSX workbook in your browser, with delimiter detection and no upload.

**Category:** Data & Files

**Free alternative to:** Convertio and Zamzar CSV to Excel (upload to a server, daily conversion limits, ads, files over 100 MB need a paid plan)

**Search phrases:** csv to excel, csv to xlsx, convert csv to excel online, open csv in excel columns, csv to xlsx utf-8, merge csv into one excel

**Allowed libraries:** papaparse, xlsx

**Published at:** https://turhe123.github.io/csv-to-excel-local/

## Acceptance checks (each must have a test)

- dropping sample-sales.csv shows a preview with the header row detected
- a semicolon separated sample is detected automatically and split into columns
- download produces a .xlsx file larger than 1 KB
- dropping two CSV files creates a workbook with two sheets
- leading zeros in a zip code column are preserved when text columns are kept as text
- works at 390px wide

## Spec

Search demand: csv to excel and csv to xlsx are evergreen queries, driven by the classic problem that double clicking a CSV in Excel puts everything in one column, or turns zip codes and long IDs into scientific notation or strips leading zeros. Online converters such as Convertio, Zamzar and CloudConvert upload the file, limit free conversions, show ads and sometimes demand an account. Business exports often hold customer data, so privacy matters.

Our version does one job: build a proper XLSX workbook from one or many CSV or TSV files. Features: drop one or several files, automatic delimiter detection with manual override (comma, semicolon, tab, pipe), a first-row-is-header toggle that turns on bold header styling and a frozen top row, a keep text columns as text toggle that stops numbers like 00123 becoming 123, an option to auto fit column widths, and a sheet per file with names taken from the file names, shortened to the 31 characters Excel allows and cleaned of forbidden characters.

UI: a drop zone above the fold. Once files load, each file appears as a card with its detected delimiter, row and column counts and a small preview of five rows the user can reorder or remove cards. Below sits an options row with defaults: header on, keep text on, auto width on. A primary Download XLSX button names the output after the first file, for example sales-2026.xlsx, or workbook.xlsx for several files. Empty state shows a short privacy line. Error state shows which file could not be parsed and why, such as an unclosed quote on a given line. Keyboard: the drop zone is a labelled file input cards have move up and move down buttons so reordering does not need dragging.

Technical approach: read each file with the File API, decode as UTF-8 and fall back to a Windows-1252 decode through TextDecoder when invalid sequences are found, strip a byte order mark, parse with papaparse using delimiter guessing, then build worksheets through the xlsx library from arrays of arrays. To preserve text, write every cell in a flagged column as a string cell type by default detect columns where any value has a leading zero or more than 15 digits and keep those as text, while converting clean numeric columns to numbers. Set column widths from the longest value capped at 60 characters. Write the workbook to an array buffer, create a Blob with the spreadsheet MIME type and download via an object URL that is revoked afterwards. Use a Web Worker for files over 5 MB.

Edge cases: Excel limits of 1,048,576 rows and 16,384 columns must produce a clear message and an offer to split into several sheets cells over 32,767 characters are truncated with a warning formula-looking values starting with equals are written as text to avoid surprises mobile Safari downloads need an anchor click from a user gesture ragged rows are padded empty files are rejected non-Latin text such as Japanese must round trip.

Demo data: a fictional CSV export from a bakery called Maple Crumb with columns order_id, zip, customer, item, qty and total, including zip codes such as 02134 and a customer named Zoë, plus a second semicolon separated file of delivery routes.
