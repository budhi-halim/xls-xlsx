# XLS to XLSX

Static, client-side spreadsheet converter for mobile and desktop browsers.

**Web app:** https://budhi-halim.github.io/xls-xlsx/

## Formats

The converter reads the file contents to detect the actual format, regardless of whether it has a misleading `.xls` extension.

- Excel 2003 SpreadsheetML XML — parsed with Python `xml.etree.ElementTree`
- HTML report saved as XLS — parsed with Python `beautifulsoup4`
- Traditional binary XLS — parsed with `xlrd`
- Modern XLSX — opened and saved with `openpyxl`

XLSX output is created with `openpyxl`. The program preserves worksheet names, cell values, numerical types, merged ranges, and applicable fonts, colors, borders, alignment, and number formats. Specialized Excel components that do not exist in the source parser (such as embedded objects and VBA macros) are not preserved.

## Browser runtime

The app uses [PyScript 2025.8.1](https://pyscript.net/) / Pyodide with `openpyxl`, `xlrd`, and `beautifulsoup4`. The browser downloads the runtime and libraries on first use; file data is processed locally and never uploaded.

Select one or more files, choose Convert, then download individual XLSX files or a batch ZIP.

## Testing

The Python engine was checked against six original company reports:
- Three Excel 2003 XML Laporan files (two worksheets each)
- Three HTML-format Rekap files (one worksheet each)

All three Rekap results matched the reference workbooks on populated cell values and merges. The three Laporan files retained both worksheets, their dimensions, and merges; this implementation also preserves several original summary cells omitted in the older reference conversions. Colors are exported as opaque ARGB to retain Excel formatting.

These are local Python-engine tests. A complete end-to-end conversion in the deployed PyScript browser UI should be checked independently before describing it as production-verified.

## Deployment

No build or server is required. Publish the repository's root folder using GitHub Pages.
