![NanoXLSX](https://raw.githubusercontent.com/rabanti-github/NanoXLSX.Compatibility/refs/heads/main/Documentation/NanoXLSXlib.png) 

# NanoXLSX.Compatibility

![NuGet Version](https://img.shields.io/nuget/v/NanoXLSX.Compatibility)
![NuGet Downloads](https://img.shields.io/nuget/dt/NanoXLSX.Compatibility)
![GitHub License](https://img.shields.io/github/license/rabanti-github/NanoXLSX.Compatibility)

NanoXLSX is a small .NET library written in C#, to create and read Microsoft Excel files in the XLSX format (Microsoft Excel 2007 or newer) in an easy and native way

---

The **Compatibility** package is responsible to add compatibility functions to NanoXLSX. Such functions may be not be used as frequently, but are especially important, when Workbooks are loaded that were created by Microsoft Excel or other Generators.

Currently supported:

- **Handling of external links**

---

Project website: [https://picoxlsx.rabanti.ch](https://picoxlsx.rabanti.ch)

See the **[Change Log](https://github.com/rabanti-github/NanoXLSX.Compatibility/blob/master/Changelog.md)** for recent updates.

## What's new in version 3.x

This is the first release if this package. It was set to v 3.x, to be consistent with the NanoXLSX v3 ecosystem

## Road Map

Possible future features (not yet in backlog):

- Preservation of workbook elements that are not implemented in NanoXLSX, when a workbook is loaded

## :robot: For AI Agents

For AI agents and LLM tooling, a machine-readable [`llms.txt`](https://raw.githubusercontent.com/rabanti-github/NanoXLSX/refs/heads/master/llms.txt) is available in the main repository.

## Important Notes about the Usage of External Links

If you add external links as:

- Part of a defined name formula
- Part of a cell formula

... keep in mind: **The formula is not checked or validated** by this package. It may lead to an **invalid Excel file**, if you add a non-compliant formula.
For example:

```C#
// The following defined name is invalid and will cause a repair dialog when opening with Excel
workbook.AddDefinedNameFormula("invalidDefinedName", "C:\\temp\\[ext.xlsx]worksheet1!$A$3+$A$4");
// The following defined name would be valid
workbook.AddDefinedNameFormula("invalidDefinedName", "C:\\temp\\[ext.xlsx]worksheet1!$A$3+C:\\temp\\[ext.xlsx]worksheet1!$A$4");
```

Furthermore, all external links, used in defined names and cells, **must be registered**, otherwise it will lead to a invalid workbook:

```C#
ExternalLink link = new ExternalLink(@"C:\temp\ext.xlsx", "ext.xlsx");
```

Last, but not least: All external links, used in formulas, are to be written **as URIs / paths** and not as identifiers like `[1]`. The URIs / paths are not checked (target file may exist or not).

## Requirements

[NanoXLSX.Compatibility](https://www.nuget.org/packages/NanoXLSX.Compatibility) is not intended as standalone package. It requires **[NanoXLSX.Core](https://www.nuget.org/packages/NanoXLSX.Core)**, and is normally part of the meta-package **[NanoXLSX](https://www.nuget.org/packages/NanoXLSX)**

**You find all technical requirements in the main repository: [NanoXLSX](https://github.com/rabanti-github/NanoXLSX)**

### General requirements

- .NET 4.5 or newer / .NET Standard
- NanoXLSX.Core as only dependency

### Utility dependencies

The Test project and GitHub Actions may also require dependencies like unit testing frameworks or workflow steps. However, **none of these dependencies are essential to build the library**. They are just utilities. The test dependencies ensure efficient unit testing and code coverage. The GitHub Actions dependencies are used for the automatization of releases and API documentation

## Installation

### Using NuGet

By package Manager (PM):

```sh
Install-Package NanoXLSX.Compatibility
```

By .NET CLI:

```sh
dotnet add package NanoXLSX.Compatibility
```

## Usage

### Quick Start (manual)

```C#
 Workbook workbook = new Workbook("worksheet1");
 // Use external link in a cell formula
 workbook.CurrentWorksheet.AddNextCellFormula("C:\\temp\\[ext1.xlsx]worksheet1!$A$1"); 

// Use external link in a defined name
 workbook.AddDefinedNameFormula("extLink1", "SUM(C:\\temp\\[ext1.xlsx]worksheet1!$A$3:$A$4)"); // within a formula expression
 workbook.AddDefinedNameFormula("extLink2", "C:\\temp\\[ext2.xlsx]worksheet1!$A$3:$A$4"); // direct reference

// Creating mandatory external links
 ExternalLink link = new ExternalLink(@"C:\temp\ext1.xlsx", "ext1.xlsx");
 ExternalLink link2 = new ExternalLink(@"C:\temp\ext2.xlsx", "ext2.xlsx");
 // Alternatively, you can use:
 // ExternalLinkBuilder builder3 = new ExternalLinkBuilder(@"C:\temp\ext1.xlsx", "ext1.xlsx");

// Add cached data
 ExternalLinkBuilder builder = link.CreateBuilder();
 builder.AddWorksheet("worksheet1");
 builder.AddCell("A1", "test");
 builder.AddCell("A2", "1", ExternalCellValue.DataType.Boolean);
 builder.AddWorksheet("ext2");
 builder.AddCell("A1", "test2");

// Add cached data
 ExternalLinkBuilder builder2 = link2.CreateBuilder();
 builder2.AddWorksheet("worksheet1");
 builder2.AddCell("A1", "test");
 builder2.AddCell("A2", "1", ExternalCellValue.DataType.Boolean);
 builder2.AddCell("A3", "10", ExternalCellValue.DataType.Number);
 builder2.AddCell("A4", "20");
 builder2.AddWorksheet("worksheet2");
 builder2.AddCell("A1", "test2");

 link2 = builder2.Build();
 link2.AddDefinedName(new ExternalDefinedName("extDefName", "A2"));

// Register external links to the workbook (otherwise the workbook is invalid)
 workbook.AddExternalLink(builder); // using a builder
 workbook.AddExternalLink(link2); // sing the final 

 workbook.SaveAs(fileName);
```

## Further References

- See the full **package API-Documentation** at: [https://rabanti-github.github.io/NanoXLSX.Compatibility/](https://rabanti-github.github.io/NanoXLSX.Compatibility/).
- See the full NanoXLSX **API-Documentation** at: [https://rabanti-github.github.io/NanoXLSX/](https://rabanti-github.github.io/NanoXLSX/).

## License

NanoXLSX.Compatibility is published under the **MIT** license.

The project / package is developed with as much compliance to this license as only possible.
Please visit the main repository [NanoXLSX](https://github.com/rabanti-github/NanoXLSX) for compliance a scan, provided by Fossa

