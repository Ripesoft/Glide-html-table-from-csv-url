# Glide Yes-Code: CSV to HTML Table

This is a sample implementation of a Yes-Code column that parses a CSV file from a URL and returns the parsed data as an array of objects.

## Getting started

To use this Yes-Code column in Glide, you need to host this code somewhere accessible via URL (e.g., on GitHub Pages, Replit, or your own server). Then, provide that URL in Glide when adding a Yes-Code column.

## Files

* `manifest.json`: Contains metadata about the function, including its name, description, parameters (isDownload, hasHeader, fileUrl), and result type.
* `function.js`: The JavaScript code that uses Papa Parse to fetch and parse the CSV from the provided URL.
* `driver.js`: Handles communication between Glide and the function.
* `index.html`: The entry point that loads the scripts.

## Parameters

- `isDownload` (boolean): Whether to download the file (usually true for URLs).
- `hasHeader` (boolean): Whether the CSV has a header row.
- `fileUrl` (string): The URL of the CSV file to parse.

## Result

Returns an array of objects representing the parsed CSV data. Each object corresponds to a row, with keys from the header (if present) or indexed numerically.

## Dependencies

This implementation uses Papa Parse for CSV parsing. Make sure Papa Parse is available (e.g., via CDN in your hosting environment).
