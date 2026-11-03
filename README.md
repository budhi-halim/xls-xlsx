# XLS to XLSX Converter

A mobile-friendly, client-side converter for legacy Excel spreadsheets.

## Use

The intended GitHub Pages URL is https://budhi-halim.github.io/xls-xlsx/ once GitHub Pages is enabled.

Choose or drag one or more files and select **Convert files**. Download results individually or as a ZIP.

## Supported formats

- Binary Excel 97–2003 `.xls` workbooks.
- Excel 2003 SpreadsheetML (XML), including XML documents saved with an `.xls` extension.
- HTML-based spreadsheet exports.
- Already-modern `.xlsx` files (validated and returned unchanged).

File type is detected by reading its contents rather than trusting the filename.

## Implementation

- Static HTML, CSS and JavaScript; no backend.
- SheetJS Community Edition 0.20.3 reads and writes spreadsheet formats within a Web Worker.
- JSZip 3.10.1 handles the optional ZIP download.
- Source file contents are never sent to an application server. The browser fetches static code and dependencies over HTTPS.

## Conversion limitations

The app performs a local round-trip check of worksheet names, nonempty cell values and formulas, and merge counts. It reports detected differences.

No converter can guarantee identical output for all legacy Excel features. Some formatting, charts, macros, images, embedded objects, external connections, and advanced features may not transfer. Review business-critical converted files before discarding originals.

## Deployment

In the repository, open **Settings → Pages**. Under **Build and deployment**, select **Deploy from a branch**, then **master** and **/(root)**. Save the settings. GitHub Pages may take a few minutes to publish.

## Local development

In **PowerShell**, set the current working directory to your local clone of `xls-xlsx`. Run:

```powershell
py -m http.server 8000
```

Visit http://localhost:8000/. The Web Worker requires an HTTP server rather than opening the HTML file with a `file://` URL.

## License

Original application code: MIT. Third-party dependencies have their own licenses.
