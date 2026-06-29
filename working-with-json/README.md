---
title: Import and Export Excel JSON in C# with Aspose.Cells
description: Import JSON into Excel cells and export worksheet ranges to JSON in C# with Aspose.Cells for .NET.
product: Aspose.Cells for .NET
category: working-with-json
language: C#
last_reviewed: 2026-06-29
---

# Import and Export Excel JSON in C# with Aspose.Cells

Use `JsonUtility.ImportData` to map JSON text into worksheet cells and `JsonUtility.ExportRangeToJson` to serialize a selected range as JSON.

| Repository fact | Value |
| --- | --- |
| Examples | 34 standalone `.cs` files |
| Primary APIs | `JsonUtility`, `JsonLayoutOptions`, `JsonSaveOptions`, `Range` |
| Excel required | No |
| Agent guidance | [`agents.md`](agents.md) |
| Catalog | [`../index.json`](../index.json) |

## Quick answer: How do I import JSON into Excel in C#?

```csharp
using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

string json = "[{\"Name\":\"Ada\",\"Quantity\":2},{\"Name\":\"Linus\",\"Quantity\":3}]";
Workbook workbook = new Workbook();
Worksheet worksheet = workbook.Worksheets[0];

JsonLayoutOptions options = new JsonLayoutOptions
{
    ArrayAsTable = true,
    ConvertNumericOrDate = true
};

JsonUtility.ImportData(json, worksheet.Cells, 0, 0, options);

if (worksheet.Cells.MaxDataRow < 1)
{
    throw new InvalidOperationException("No JSON rows were imported.");
}

workbook.Save("json-import.xlsx");
Console.WriteLine("Imported JSON into json-import.xlsx.");
```

## API choice

| Need | API |
| --- | --- |
| JSON text to cells | `JsonUtility.ImportData` |
| Control import layout | `JsonLayoutOptions` |
| Selected cells to JSON | `JsonUtility.ExportRangeToJson` |
| Control export | `JsonSaveOptions` |

## Featured examples

- [Export a worksheet range to JSON](export-a-specific-cell-range-from-a-worksheet-to-json-using-exportrangetojsonoptions.cs)
- [Map multiple JSON tables to worksheets](load-a-json-array-representing-multiple-tables-and-map-each-element-to-a-separate-worksheet.cs)
- [Handle malformed JSON input](implement-error-handling-to-catch-jsonutilityload-exceptions-when-source-json-file-is-malformed.cs)
- [Merge multiple JSON files into one workbook](merge-multiple-json-files-into-a-single-workbook-creating-separate-worksheets-for-each-file.cs)
- [Print worksheet JSON to standard output](create-a-console-application-that-prints-json-representation-of-the-first-worksheet-to-standard-output.cs)
- [Validate exported JSON row count](validate-that-exported-json-includes-expected-number-of-rows-by-comparing-with-original-worksheet-row-count.cs)

These generated candidates may contain obsolete or inferred APIs. In particular, verify current export option types and signatures before reuse.

## Getting started

```bash
dotnet new console -n ExcelJsonExample
cd ExcelJsonExample
dotnet add package Aspose.Cells
dotnet build
dotnet run
```

Import offsets are zero-based. Export only the required `Range`; workbook-wide or hidden data should not be exposed accidentally. Parse exported JSON to verify keys, values, types, and row counts.

## FAQ

**Can Aspose.Cells import JSON without Excel?** Yes.

**How do I export only part of a worksheet?** Create a `Range` and pass it to `JsonUtility.ExportRangeToJson`.

**Is Aspose.Cells a JSON schema validator?** No. Use a JSON validation library for that separate concern.

**Which export options should I use?** Prefer the current `JsonSaveOptions` overload verified against the installed package.

**How should untrusted JSON be handled?** Limit size, depth, row/column expansion, and output scope; do not log sensitive content.

## Related and official resources

- [`cells-data`](../cells-data/)
- [`working-with-tables`](../working-with-tables/)
- [`xml-maps`](../xml-maps/)
- [JsonUtility](https://reference.aspose.com/cells/net/aspose.cells.utility/jsonutility/)
- [JsonLayoutOptions](https://reference.aspose.com/cells/net/aspose.cells.utility/jsonlayoutoptions/)
- [JsonSaveOptions](https://reference.aspose.com/cells/net/aspose.cells/jsonsaveoptions/)
- [Aspose.Cells NuGet](https://www.nuget.org/packages/Aspose.Cells/)

Use an Aspose.Cells license or temporary license for full evaluation and review the repository [`LICENSE`](../LICENSE).
