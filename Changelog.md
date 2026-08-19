# Change Log - NanoXLSX.Compatibility

## v3.2.0

---
Release Date: **20.08.2026** <sup>(DMY)</sup>

- Initial release of the compatibility library
- Read and write support
- Compatibility support and resolution of external links in formulas or defined names
- Adds several accessor methods to the Workbook class
- Exposing of the public data types:
  - `ExternalLink` (Contains link and caching information)
  - `ExternalLinkBuilder` (Builder to create external links with worksheets, cells and defined names)
  - `ExternalWorksheet` (Contains cached worksheet data of external links)
  - `ExternalCellValue` (Contains cached cell data in worksheets of external links)
  - `ExternalDefinedName` (Contains defined names of external links)

Note *I*: The version of the package is set to 3.2.0  instead 1.0.0, to be consistent with the NanoXLSX v3 ecosystem
