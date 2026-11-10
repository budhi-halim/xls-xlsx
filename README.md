# XLS to XLSX Converter

A client-side spreadsheet converter for mobile and desktop browsers.

**Application:** https://budhi-halim.github.io/xls-xlsx/

## Supported files

| Input | Handling |
| --- | --- |
| Excel 2003 SpreadsheetML (.xls, .xml) | Dedicated XML parser; retains worksheet names, cell contents, merged ranges, number formats, borders, colors, and alignment |
| HTML table exported as .xls | Dedicated HTML parser; preserves table structure, spans, styled sections, and numeric Rupiah amounts |
| Excel 97–2003 binary .xls | SheetJS Community Edition 0.20.3 fallback |
| .xlsx | Returned unchanged |

The converter detects the underlying file format from its contents, not its extension. XML and HTML files are processed without external dependencies. Binary XLS uses a static SheetJS library fetched on demand. ZIP download uses JSZip from a CDN; individual XLSX downloads do not need JSZip.

All processing is performed in the user's browser. Source spreadsheet contents are not uploaded.

## Validation

The custom XML and HTML converter was tested against six original .xls samples and their separately generated Python-based XLSX references (three Laporan and three Rekap). For all three Rekap samples, every populated cell and merged range matched the reference. For Laporan, the source cell data, worksheet names, and merged ranges were preserved. The reference conversion had placed some final-row total values at different cell addresses; this converter retains their positions from the source XML.

Tests covered browser conversion, single-file downloads, simultaneous Laporan/Rekap selection, and mobile-sized layouts. Converted files were independently imported with a spreadsheet library to check workbook validity. The six private reports are not included in this public repository.

## Local development

In PowerShell, with the repository folder as the working directory:

```powershell
py -m http.server 8000
```

Open http://localhost:8000/. No build step or backend is required.

## Limitations

No guarantee of perfect preservation for charts, macros, embedded objects, and advanced features in arbitrary legacy files. In particular, the binary XLS fallback may not preserve all visual formatting. Review business-critical converted workbooks before replacing the originals.
