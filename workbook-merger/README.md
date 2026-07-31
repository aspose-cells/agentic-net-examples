---
title: Merge Excel Workbooks in C# with Aspose.Cells for .NET
description: Combine XLS and XLSX workbooks or copy selected worksheets between files in C# without Microsoft Excel.
product: Aspose.Cells for .NET
category: workbook-merger
language: C#
last_reviewed: 2026-06-29
---

# Merge Excel Workbooks in C# with Aspose.Cells for .NET

Use `Workbook.Combine` to append complete workbooks and `Worksheet.Copy` to copy selected sheets into another workbook without Microsoft Excel.

| Repository fact | Value |
| --- | --- |
| Examples | 30 standalone `.cs` files |
| Primary APIs | `Workbook.Combine`, `Worksheet.Copy`, `Workbook.Save` |
| Excel required | No |
| Agent guidance | [`agents.md`](agents.md) |
| Catalog | [`../index.json`](../index.json) |

## Quick answer: How do I merge Excel files in C#?

```csharp
using System;
using Aspose.Cells;

Workbook destination = new Workbook();
destination.Worksheets[0].Name = "North";
destination.Worksheets[0].Cells["A1"].PutValue("North");

Workbook source = new Workbook();
source.Worksheets[0].Name = "South";
source.Worksheets[0].Cells["A1"].PutValue("South");

destination.Combine(source);

if (destination.Worksheets.Count != 2)
{
    throw new InvalidOperationException("Merge failed.");
}

destination.Save("merged-workbooks.xlsx");
Console.WriteLine("Created merged-workbooks.xlsx with 2 worksheets.");
```

## API choice

| Need | API |
| --- | --- |
| Merge complete workbooks | `Workbook.Combine` |
| Copy selected sheet between workbooks | `Worksheet.Copy` |
| Copy within one workbook | `WorksheetCollection.AddCopy` |
| Refresh merged formulas | `Workbook.CalculateFormula` |
| Output file or stream | `Workbook.Save` |

## Featured examples

- [Combine moderate-size workbooks](use-workbookcombine-to-merge-two-or-more-workbooks-when-file-sizes-are-moderate.cs)
- [Copy selected worksheets](copy-specific-worksheets-from-source-workbooks-into-the-target-workbook-using-worksheetcopy-method.cs)
- [Verify merged worksheet count](verify-that-the-merged-workbook-contains-the-expected-number-of-worksheets-after-combination.cs)
- [Save a merged workbook to memory](save-the-merged-workbook-into-a-memory-stream-for-immediate-transmission-via-a-web-api.cs)
- [Check embedded images after merge](confirm-that-embedded-images-are-retained-in-the-merged-workbook-after-using-workbookcombine.cs)
- [Check charts after merge](check-that-all-charts-from-source-workbooks-appear-correctly-in-the-combined-workbook.cs)

Generated examples may depend on files or contain version-sensitive APIs. Validate each one before use.

## Getting started

```bash
dotnet new console -n WorkbookMerger
cd WorkbookMerger
dotnet add package Aspose.Cells
dotnet build
dotnet run
```

`Combine` changes the destination workbook. Plan for its initial sheet, resolve duplicate names, and reopen the saved output to verify formulas, styles, links, charts, images, and macros.

## FAQ

**Combine or Copy?** Use `Combine` for all workbook content and `Worksheet.Copy` for selected sheets.

**Why is there a blank sheet?** A new `Workbook` starts with a default worksheet.

**Are formulas recalculated automatically?** Call `CalculateFormula` when current results are required.

**Can macro workbooks be merged?** Use a macro-compatible output and verify the exact package/version behavior.

**Can I merge to a stream?** Yes; reset stream position before a downstream consumer reads it.

## Related and official resources

- [`working-with-worksheets`](../working-with-worksheets/)
- [`calculate-formulas`](../calculate-formulas/)
- [`save-workbook`](../save-workbook/)
- [Combine Workbooks](https://docs.aspose.com/cells/net/combining-multiple-workbooks-into-a-single-workbook/)
- [Workbook.Combine](https://reference.aspose.com/cells/net/aspose.cells/workbook/combine/)
- [Worksheet.Copy](https://reference.aspose.com/cells/net/aspose.cells/worksheet/copy/)
- [Aspose.Cells NuGet](https://www.nuget.org/packages/Aspose.Cells/)

Review the repository [`LICENSE`](../LICENSE) and Aspose licensing terms before production use.
