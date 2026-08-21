---
title: Work with Excel Worksheets in C# using Aspose.Cells for .NET
description: C# examples for adding, accessing, copying, moving, hiding, freezing, protecting, and configuring Excel worksheets without Microsoft Excel.
product: Aspose.Cells for .NET
category: working-with-worksheets
language: C#
last_reviewed: 2026-08-14
---

# Work with Excel Worksheets in C# using Aspose.Cells for .NET

Create, access, copy, move, rename, hide, freeze, protect, and configure Microsoft Excel worksheets in C# with Aspose.Cells for .NET, without installing Microsoft Excel. This category covers worksheet lifecycle, tab organization, pane settings, page layout, printing, and protected editing.

The main entry point is [`Workbook.Worksheets`](https://reference.aspose.com/cells/net/aspose.cells/workbook/worksheets/), which returns the workbook's `WorksheetCollection`.

| Repository fact | Value |
| --- | --- |
| Product | Aspose.Cells for .NET |
| Language | C# |
| Category | Worksheet operations |
| Examples | 489 standalone `.cs` files |
| Primary APIs | `Workbook.Worksheets`, `WorksheetCollection`, `Worksheet` |
| Other key APIs | `Cells`, `PageSetup`, `Worksheet.FreezePanes`, `Worksheet.Protect` |
| Microsoft Excel required | No |
| Agent instructions | [`agents.md`](agents.md) |
| Machine-readable catalog | [`../index.json`](../index.json) |

## Quick answer: How do I add an Excel worksheet in C#?

Create a workbook, call `Workbook.Worksheets.Add`, populate the returned worksheet, and save the workbook.

```csharp
using System;
using Aspose.Cells;

namespace AddExcelWorksheet
{
    internal class Program
    {
        static void Main()
        {
            Workbook workbook = new Workbook();

            Worksheet worksheet = workbook.Worksheets.Add("Sales");
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Revenue");
            worksheet.Cells["A2"].PutValue("Widget");
            worksheet.Cells["B2"].PutValue(1250);

            worksheet.FreezePanes(1, 0, 1, 0);
            worksheet.AutoFitColumns();

            Console.WriteLine($"Added worksheet: {worksheet.Name}");
            Console.WriteLine($"Worksheet count: {workbook.Worksheets.Count}");

            workbook.Save("worksheet-result.xlsx");
        }
    }
}
```

Expected result:

```text
Added worksheet: Sales
Worksheet count: 2
```

## What this category covers

Use these examples to answer questions such as:

- How do I add a named worksheet to an Excel workbook?
- How do I access a worksheet by name or zero-based index?
- How do I rename, copy, move, hide, unhide, or remove a sheet?
- How do I freeze a header row or identifier columns?
- How do I protect a worksheet while keeping input cells editable?
- How do I configure page orientation, margins, headers, and footers?
- How do I set print areas and manual page breaks?
- How do I insert or delete worksheet rows and columns?
- How do I change worksheet tab color, zoom, or gridline display?
- How do I verify worksheet order, visibility, and persisted settings?

## Choose the right worksheet API

| Developer goal | API | Notes |
| --- | --- | --- |
| Add a named worksheet | `workbook.Worksheets.Add("Name")` | Returns the new `Worksheet` |
| Access by index | `workbook.Worksheets[0]` | Indexes are zero-based |
| Access by name | `workbook.Worksheets["Name"]` | Check the result when the name is not guaranteed |
| Copy within a workbook | `workbook.Worksheets.AddCopy(indexOrName)` | Returns the new sheet index |
| Copy between worksheets | `destination.Copy(source)` | Verify formulas, objects, and page settings |
| Reorder a sheet | `worksheet.MoveTo(index)` | Changes tab order, not cell data |
| Hide or show a sheet | `worksheet.IsVisible` | Keep at least one sheet visible |
| Freeze panes | `worksheet.FreezePanes(...)` | Use zero-based boundaries deliberately |
| Protect a sheet | `worksheet.Protect(...)` | Unlock intended input cells first |
| Configure printing | `worksheet.PageSetup` | Settings are per worksheet |

## Featured worksheet examples

### Access, add, rename, and organize sheets

- [Access a worksheet by name](access-a-worksheet-by-its-name-and-store-the-reference-for-further-operations.cs)
- [Access a worksheet by zero-based index](access-a-worksheet-by-its-zerobased-index-and-assign-it-to-a-variable.cs)
- [Rename the active worksheet while preserving data](rename-the-active-worksheet-to-quarterlyreport-while-preserving-all-existing-cell-data.cs)
- [Move a worksheet to a new tab position](move-a-worksheet-to-a-new-position-by-providing-the-target-index-within-the-same-workbook.cs)

### Copy worksheets

- [Copy a worksheet within the same workbook by name](copy-a-worksheet-within-the-same-workbook-using-its-name-and-ensure-content-integrity.cs)
- [Copy a worksheet within the same workbook by index](copy-a-worksheet-within-the-same-workbook-using-its-numeric-index-and-verify-duplication.cs)
- [Copy a worksheet between workbooks and preserve formulas](copy-a-worksheet-from-a-source-workbook-to-a-target-workbook-while-preserving-formulas.cs)
- [Copy multiple worksheets with a name prefix](copy-multiple-worksheets-whose-names-start-with-a-specific-prefix-into-a-new-workbook-for-backup.cs)

### Freeze panes and worksheet views

- [Freeze the first row and first column](freeze-panes-at-row-one-and-column-one-to-keep-headers-visible-while-scrolling.cs)
- [Freeze the top row and first two columns](freeze-the-top-row-and-first-two-columns-to-keep-headers-visible-during-scrolling.cs)
- [Freeze the leftmost two columns](freeze-the-leftmost-two-columns-by-calling-freezepanes-with-row-zero-and-column-two.cs)
- [Autofit columns before freezing panes](autofit-all-columns-before-freezing-to-preserve-column-widths-after-view-changes.cs)

### Protect and configure worksheet editing

- [Protect a worksheet while allowing edits in unlocked cells](protect-the-worksheet-with-a-password-and-allow-users-to-edit-only-unlocked-cells.cs)
- [Protect individual cells while leaving other cells editable](protect-individual-cells-with-a-password-while-leaving-other-cells-editable-for-users.cs)
- [Protect a worksheet after freezing header rows](protect-the-worksheet-after-freezing-rows-to-prevent-accidental-changes-to-the-header-area.cs)

### Rows, columns, and print layout

- [Insert five rows and shift existing data](insert-five-new-rows-at-position-ten-shifting-existing-rows-downward-accordingly.cs)
- [Delete rows 20 through 30](delete-rows-twenty-through-thirty-from-the-worksheet-and-adjust-formulas-accordingly.cs)
- [Insert a new column at index five](insert-a-new-column-at-index-five-and-shift-existing-columns-to-the-right.cs)
- [Copy page setup between worksheets](copy-the-entire-pagesetup-configuration-from-a-source-worksheet-to-a-destination-worksheet.cs)
- [Insert a manual page break after row 30](insert-a-manual-page-break-after-row-30-to-control-pagination-in-the-printed-document.cs)

> This repository contains generated examples across broad and sometimes version-sensitive scenarios. Confirm APIs against the installed package and follow [`agents.md`](agents.md) before using an example in production.

## Getting started

### Prerequisites

- A supported .NET SDK
- The `Aspose.Cells` NuGet package
- An Aspose.Cells license for production use or a temporary license for full evaluation

### Install Aspose.Cells

```bash
dotnet new console -n WorksheetExample
cd WorksheetExample
dotnet add package Aspose.Cells
```

Copy an example into `Program.cs`, then run:

```bash
dotnet build
dotnet run
```

## Worksheet fundamentals

### Access worksheets by index or name

```csharp
Worksheet firstSheet = workbook.Worksheets[0];
Worksheet salesSheet = workbook.Worksheets["Sales"];
```

Indexes are zero-based. Use a name when the workbook contract guarantees the name, and guard dynamic indexes and names.

### Copy a worksheet within the workbook

```csharp
int copiedIndex = workbook.Worksheets.AddCopy("Sales");
Worksheet copiedSheet = workbook.Worksheets[copiedIndex];
copiedSheet.Name = "Sales Copy";
```

### Move a worksheet

```csharp
Worksheet worksheet = workbook.Worksheets["Sales"];
worksheet.MoveTo(0);
```

Moving changes tab order; it does not delete the sheet's content.

### Freeze a header row

```csharp
worksheet.FreezePanes(1, 0, 1, 0);
```

This freezes the one row above zero-based row index `1`.

### Protect a worksheet with editable input cells

```csharp
Style unlockedStyle = workbook.CreateStyle();
unlockedStyle.IsLocked = false;
worksheet.Cells["B2"].SetStyle(unlockedStyle);

worksheet.Protect(ProtectionType.All, "change-me", null);
```

Worksheet protection controls editing behavior. It does not encrypt the workbook.

### Configure print layout

```csharp
PageSetup pageSetup = worksheet.PageSetup;
pageSetup.Orientation = PageOrientationType.Landscape;
pageSetup.FitToPagesWide = 1;
pageSetup.FitToPagesTall = 0;
```

Render or inspect output when page-layout fidelity is the actual requirement.

## Worksheet FAQ

### Can Aspose.Cells manage worksheets without Microsoft Excel?

Yes. Aspose.Cells creates and modifies worksheet collections directly without Excel, Office automation, or Interop.

### Are worksheet indexes zero-based?

Yes. `workbook.Worksheets[0]` is the first worksheet.

### How do I duplicate a worksheet?

Use `WorksheetCollection.AddCopy` to copy a sheet within the same workbook. Use `Worksheet.Copy` when copying source worksheet content into another worksheet, including cross-workbook workflows supported by the installed version.

### How do I keep the top row visible?

Call `worksheet.FreezePanes(1, 0, 1, 0)`.

### Can I hide every worksheet?

Do not do so. Keep at least one worksheet visible and active so the saved workbook remains valid and usable.

### Does worksheet protection encrypt the file?

No. Protection limits editing actions. Use workbook encryption when confidentiality is required.

### Does renaming or moving a worksheet change its data?

No. `Worksheet.Name` changes the tab name and `Worksheet.MoveTo` changes tab order. Verify formulas and references when a larger workflow depends on names or positions.

### When should I use `PageSetup`?

Use it for worksheet-specific print settings such as orientation, paper size, margins, headers, footers, print areas, titles, and scaling.

## Guidance for AI coding agents and RAG systems

For reliable worksheet answers:

1. Match the request to one worksheet lifecycle or configuration operation.
2. Use `Workbook.Worksheets` as the collection entry point.
3. Guard dynamic indexes and names.
4. Keep one worksheet visible.
5. Verify the resulting property, order, count, or persisted workbook.
6. Cite the relevant example or official API when attribution is required.

Useful retrieval aliases:

- add Excel worksheet in C#
- access worksheet by name Aspose.Cells
- copy Excel sheet without Interop
- move or reorder worksheet tabs
- freeze top row in Excel using .NET
- protect sheet and unlock cells
- configure Excel worksheet page setup

## Related categories

- [`manage-workbook`](../manage-workbook/) - workbook-level structure and settings
- [`rows-and-columns`](../rows-and-columns/) - detailed row and column operations
- [`cells-data`](../cells-data/) - read, write, import, and export cell data
- [`format-cells`](../format-cells/) - cell and range formatting
- [`working-with-tables`](../working-with-tables/) - structured tables
- [`working-with-charts`](../working-with-charts/) - worksheet charts
- [`working-with-shapes`](../working-with-shapes/) - drawing objects and controls

## Official Aspose.Cells resources

- [Manage Worksheets documentation](https://docs.aspose.com/cells/net/managing-worksheets/)
- [Worksheet API](https://reference.aspose.com/cells/net/aspose.cells/worksheet/)
- [WorksheetCollection API](https://reference.aspose.com/cells/net/aspose.cells/worksheetcollection/)
- [PageSetup API](https://reference.aspose.com/cells/net/aspose.cells/pagesetup/)
- [Protect Worksheets documentation](https://docs.aspose.com/cells/net/protecting-worksheets/)
- [Aspose.Cells NuGet package](https://www.nuget.org/packages/Aspose.Cells/)

## Validation and trust

Validate every reused example with the repository's installed Aspose.Cells version and target framework. Confirm collection state in memory, save the workbook, and reopen it when persistence matters. For page setup and view operations, inspect the resulting workbook or rendered output rather than relying only on file creation.

## License

These examples use [Aspose.Cells for .NET](https://products.aspose.com/cells/net/). Review the repository [`LICENSE`](../LICENSE) and Aspose licensing terms before production use.
