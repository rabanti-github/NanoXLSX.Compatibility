![NanoXLSX](https://raw.githubusercontent.com/rabanti-github/NanoXLSX.Compatibility/refs/heads/main/Documentation/NanoXLSXlib.png) 

# NanoXLSX.Compatibility

![NuGet Version](https://img.shields.io/nuget/v/NanoXLSX.Compatibility)
![NuGet Downloads](https://img.shields.io/nuget/dt/NanoXLSX.Compatibility)
![GitHub License](https://img.shields.io/github/license/rabanti-github/NanoXLSX.Compatibility)

NanoXLSX is a small .NET library written in C#, to create and read Microsoft Excel files in the XLSX format (Microsoft Excel 2007 or newer) in an easy and native way

---

The **Compatibility** package is responsible to add compatibility functions to NanoXLSX. Such functions may be not be used as frequently, but are especially important, when Workbooks are loaded that were created by Microsoft Excel or other Generators.

Currently supported:

- Handling of external links

---

Project website: [https://picoxlsx.rabanti.ch](https://picoxlsx.rabanti.ch)

See the **[Change Log](https://github.com/rabanti-github/NanoXLSX.Compatibility/blob/master/Changelog.md)** for recent updates.

## What's new in version 3.x

This is the first release if this package. It was set to v 3.x, to be consistent with the NanoXLSX v3 ecosystem

## Road map

Possible future features (not yet in backlog):

- Preservation of workbook elements that are not implemented in NanoXLSX, when a workbook is loaded

## :robot: For AI Agents

For AI agents and LLM tooling, a machine-readable [`llms.txt`](https://raw.githubusercontent.com/rabanti-github/NanoXLSX/refs/heads/master/llms.txt) is available in the main repository.

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

[TBD]

## Further References

- See the full **package API-Documentation** at: [https://rabanti-github.github.io/NanoXLSX.Compatibility/](https://rabanti-github.github.io/NanoXLSX.Compatibility/).
- See the full NanoXLSX **API-Documentation** at: [https://rabanti-github.github.io/NanoXLSX/](https://rabanti-github.github.io/NanoXLSX/).

## License

NanoXLSX.Compatibility is published under the **MIT** license.

The project / package is developed with as much compliance to this license as only possible.
Please visit the main repository [NanoXLSX](https://github.com/rabanti-github/NanoXLSX) for compliance a scan, provided by Fossa 
