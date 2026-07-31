---
title: Create and Manage Excel Tables in C# with Aspose.Cells
description: Create, name, style, filter, resize, total, and convert structured Excel tables in C# without Microsoft Excel.
product: Aspose.Cells for .NET
category: working-with-tables
language: C#
last_reviewed: 2026-06-29
---

# Create and Manage Excel Tables in C# with Aspose.Cells

Create and manage structured Excel tables (`ListObject`) with `Worksheet.ListObjects`, without Microsoft Excel.

| Repository fact | Value |
| --- | --- |
| Examples | 143 standalone `.cs` files |
| Primary APIs | `ListObjectCollection`, `ListObject`, `ListColumn`, `TableStyleType` |
| Excel required | No |
| Agent guidance | [`agents.md`](agents.md) |
| Catalog | [`../index.json`](../index.json) |

## Quick answer: How do I create an Excel table in C#?

```csharp
using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

Workbook workbook = new Workbook();
Worksheet worksheet = workbook.Worksheets[0];
worksheet.Cells["A1"].PutValue("Product");
worksheet.Cells["B1"].PutValue("Quantity");
worksheet.Cells["A2"].PutValue("Notebook");
worksheet.Cells["B2"].PutValue(12);
worksheet.Cells["A3"].PutValue("Pen");
worksheet.Cells["B3"].PutValue(30);

int index = worksheet.ListObjects.Add("A1", "B3", true);
ListObject table = worksheet.ListObjects[index];
table.DisplayName = "InventoryTable";
table.TableStyleType = TableStyleType.TableStyleMedium2;

if (worksheet.ListObjects.Count != 1)
{
    throw new InvalidOperationException("Table creation failed.");
}

workbook.Save("excel-table.xlsx");
Console.WriteLine("Created InventoryTable with two data rows.");
```

## API choice

| Need | API |
| --- | --- |
| Create or enumerate tables | `Worksheet.ListObjects` |
| Name, resize, style, totals | `ListObject` |
| Work with table columns | `ListColumn` |
| Built-in visual style | `TableStyleType` |
| Filter rows | `ListObject.AutoFilter` |
| Remove table semantics | `ConvertToRange` |

## Featured examples

- [Create and name a table from a range](create-a-new-worksheet-table-from-a-range-of-cells-and-assign-a-custom-name.cs)
- [Add a totals row with sum formulas](add-a-totals-row-to-the-table-and-configure-sum-formulas-for-numeric-columns.cs)
- [Apply a built-in table style](apply-a-builtin-table-style-that-matches-the-workbooks-theme-for-consistent-visual-appearance.cs)
- [Enable AutoFilter and filter above a threshold](enable-autofilter-on-the-table-and-define-a-filter-to-show-only-rows-with-values-above-threshold.cs)
- [Convert a table to a range and copy it](convert-a-table-to-a-range-and-copy-the-resulting-range-to-another-worksheet-using-rangecopy.cs)
- [Update a table cell after locating its table](overwrite-an-existing-value-in-a-table-cell-using-cellputvalue-after-retrieving-the-table-with-cellgettable.cs)

Generated examples can be specialized or version-sensitive. Compile, run, and verify table semantics before reuse.

## Getting started

```bash
dotnet new console -n ExcelTableExample
cd ExcelTableExample
dotnet add package Aspose.Cells
dotnet build
dotnet run
```

The table range is inclusive. The final `true` argument means the first row already contains headers. Use unique valid headers and a workbook-unique `DisplayName`.

## FAQ

**Is a styled range an Excel table?** No. A real table is a `ListObject`.

**Does filtering delete rows?** No. It changes row visibility.

**What happens after `ConvertToRange`?** Table semantics are removed; formatting may remain.

**When should formulas be calculated?** Before reading calculated-column or totals results after source data changes.

## Related and official resources

- [`cells-data`](../cells-data/)
- [`managing-ranges`](../managing-ranges/)
- [`calculate-formulas`](../calculate-formulas/)
- [`slicer`](../slicer/)
- [Create and Manage Tables](https://docs.aspose.com/cells/net/create-and-manage-table/)
- [ListObject](https://reference.aspose.com/cells/net/aspose.cells.tables/listobject/)
- [ListObjectCollection](https://reference.aspose.com/cells/net/aspose.cells.tables/listobjectcollection/)
- [Aspose.Cells NuGet](https://www.nuget.org/packages/Aspose.Cells/)

Review the repository [`LICENSE`](../LICENSE) and Aspose licensing terms before production use.
