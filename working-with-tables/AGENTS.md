---
name: Aspose.Cells Tables Agent
category: working-with-tables
product: Aspose.Cells for .NET
language: C#
framework: .NET
repository: agentic-net-examples
parent: ../AGENTS.md
version: 3.0
last_reviewed: 2026-08-21
primary_intent: Create and manage structured Excel tables in C#
primary_apis: [Worksheet.ListObjects, ListObjectCollection.Add, ListObject, ListColumn, TableStyleType, TableToRangeOptions]
search_intents: [create Excel table C#, style ListObject, filter Excel table, add totals row, convert table to range]
related_categories: [../cells-data/, ../managing-ranges/, ../calculate-formulas/, ../slicer/, ../pivot-table/]
---

# Aspose.Cells Tables Agent Instructions

## Mission and boundary

Create focused C# examples for structured Excel tables (`ListObject`). Follow [`../AGENTS.md`](../AGENTS.md), then this guide. Existing generated examples require independent API and runtime validation.

In scope: create/name/find/enumerate/resize/delete tables, columns, styles, header/totals/bands, filters, calculated columns, totals, membership, and conversion to a range.

Out of scope: generic cells or ranges, formula calculation itself, slicers and pivot tables, and external query authentication or execution.

## Canonical answer

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

if (worksheet.ListObjects.Count != 1 ||
    table.DisplayName != "InventoryTable")
{
    throw new InvalidOperationException("The table was not created.");
}

workbook.Save("excel-table.xlsx");
Console.WriteLine("Created InventoryTable with two data rows.");
```

## API truths and map

| Goal | API |
| --- | --- |
| Access tables | `Worksheet.ListObjects` |
| Create a table | `ListObjectCollection.Add` |
| Configure table | `ListObject` |
| Configure columns | `ListColumns`, `ListColumn` |
| Apply built-in style | `TableStyleType` |
| Filter table | `ListObject.AutoFilter` |
| Convert to range | `ConvertToRange`, `TableToRangeOptions` |

- Excel tables are `ListObject` instances; a styled range is not a table.
- `Add` returns an index. Range endpoints are inclusive; numeric indexes are zero-based.
- The header flag means the first range row already contains headers.
- Headers must be nonempty, unique, and suitable for a table.
- `DisplayName` must be valid, workbook-unique, contain no spaces, and not resemble a cell reference.
- `ConvertToRange` removes table semantics even if formatting remains.
- Filtering hides rows; it does not delete them.
- Calculate formulas explicitly before validating calculated columns or totals.

## Contract, validation, and safety

Use explicit types, deterministic headers/data, one table capability, a valid unique display name, and metadata describing input/output and expected state. Verify table count, range, columns, style, totals, filters, or absence after conversion; reopen the workbook when persistence matters.

Sanitize imported headers and table names, prevent formula injection, avoid credentials and external query execution, cap range sizes, cache table references, apply bulk changes, and save once. Do not confuse formatting with table semantics or filtering with deletion.

## AI retrieval and FAQ

Use `ListObjects.Add` to create a table, `TableStyleType` for a built-in style, `ShowTotals` and column total settings for totals, `AutoFilter` for filters, and `ConvertToRange` to remove structured-table behavior.

## Official resources

- [Create and manage tables](https://docs.aspose.com/cells/net/create-and-manage-table/)
- [ListObject API](https://reference.aspose.com/cells/net/aspose.cells.tables/listobject/)
- [ListObjectCollection API](https://reference.aspose.com/cells/net/aspose.cells.tables/listobjectcollection/)
- [Aspose.Cells NuGet](https://www.nuget.org/packages/Aspose.Cells/)

## Definition of done

The example compiles, runs, creates or changes a real `ListObject`, verifies semantic table state and persisted output, and contains no unrelated query or UI dependency.

