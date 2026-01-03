# Glide Yes-Code: CSV/XLS to Data Array

This is a sample implementation of a Yes-Code column that parses CSV, TSV, XLS, or XLSX files from a URL and returns the parsed data as an array of objects or arrays.

## Getting started

To use this Yes-Code column in Glide, you need to host this code somewhere accessible via URL (e.g., on GitHub Pages, Replit, or your own server). Then, provide that URL in Glide when adding a Yes-Code column.

## Files

* `manifest.json`: Contains metadata about the function, including its name, description, parameters (isDownload, hasHeader, fileUrl), and result type.
* `function.js`: The JavaScript code that uses Papa Parse for CSV/TSV and SheetJS for XLS/XLSX files.
* `driver.js`: Handles communication between Glide and the function.
* `index.html`: The entry point that loads the scripts.

## Parameters

- `isDownload` (boolean): Whether to download the file (used for CSV/TSV; XLS/XLSX always download).
- `hasHeader` (boolean): Whether the file has a header row.
- `fileUrl` (string): The URL of the file to parse (supports .csv, .tsv, .xls, .xlsx).

## Result

Returns an array of parsed data. If `hasHeader` is true, returns an array of objects with keys from the header. If false, returns an array of arrays.

## Dependencies

This implementation uses Papa Parse for CSV/TSV parsing and SheetJS (XLSX) for Excel file parsing. Both are included locally.
