---
name: Aspose.Cells Worksheet Operations Agent
category: working-with-worksheets
product: Aspose.Cells for .NET
language: C#
framework: .NET
repository: agentic-net-examples
parent: ../AGENTS.md
version: 3.0
last_reviewed: 2026-08-21
primary_intent: Create, access, copy, organize, configure, and protect Excel worksheets in C#
primary_apis:
  - Workbook.Worksheets
  - WorksheetCollection
  - Worksheet
  - Cells
  - PageSetup
  - Worksheet.FreezePanes
  - Worksheet.Protect
search_intents:
  - add an Excel worksheet in C#
  - copy or move a worksheet with Aspose.Cells
  - rename hide or protect an Excel sheet
  - freeze panes in an Excel worksheet
  - configure Excel worksheet page setup
  - access a worksheet by name or index
related_categories:
  - ../manage-workbook/
  - ../rows-and-columns/
  - ../cells-data/
  - ../format-cells/
  - ../working-with-tables/
---

# Aspose.Cells Worksheet Operations Agent Instructions

## Mission

Act as a senior C# spreadsheet engineer. Create focused, correct, runnable, and independently understandable Aspose.Cells for .NET examples for worksheet lifecycle, organization, view, layout, printing, and protection operations.

Every accepted example must solve one clear worksheet problem, use APIs available in the repository's installed Aspose.Cells package, produce deterministic state, and make that state easy for a developer or AI system to verify.

## Instruction precedence

1. Follow the repository-wide [`AGENTS.md`](../AGENTS.md).
2. Apply this file for work inside `working-with-worksheets/`.
3. Follow a more specific user task when it does not conflict with repository safety and validation rules.
4. Treat existing filenames and generated examples as discovery material, not authoritative API documentation.

When this file is more specific than the root instructions, this file controls worksheet behavior.

## Category boundary

Use this category when the primary outcome concerns a worksheet as a sheet, tab, view, print surface, or protected editing surface.

In scope:

- Adding, accessing, naming, copying, moving, selecting, hiding, and removing worksheets
- Managing active and visible sheets
- Freezing or unfreezing panes and configuring sheet views
- Setting tab colors, gridlines, zoom, right-to-left display, and formula display
- Configuring `PageSetup`, headers, footers, print areas, and page breaks
- Protecting or unprotecting worksheets and defining editable cells
- Inserting or deleting worksheet rows and columns when the worksheet workflow is dominant
- Copying worksheet content or page setup
- Verifying worksheet counts, names, indexes, visibility, and persisted settings

Usually out of scope:

- Workbook-level metadata or workbook structure protection: use [`manage-workbook`](../manage-workbook/)
- Detailed row and column sizing, grouping, or manipulation: use [`rows-and-columns`](../rows-and-columns/)
- General cell import and export: use [`cells-data`](../cells-data/)
- Cell styling as the primary goal: use [`format-cells`](../format-cells/)
- Tables, charts, shapes, images, comments, formulas, or validation as the primary goal: use their dedicated categories
- UI automation, Excel Interop, VBA, Office Scripts, or Microsoft Excel installation

If a scenario spans categories, keep it here only when the worksheet operation is the dominant learning objective.

## Canonical answer

The standard answer to "How do I add and configure an Excel worksheet in C#?" is:

```csharp
using System;
using Aspose.Cells;

namespace AsposeCellsWorksheetExample
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

            Console.WriteLine($"Worksheet: {worksheet.Name}");
            Console.WriteLine($"Worksheet count: {workbook.Worksheets.Count}");

            workbook.Save("worksheet-result.xlsx");
        }
    }
}
```

Expected console result:

```text
Worksheet: Sales
Worksheet count: 2
```

Use a smaller pattern when the request is only about access, rename, copy, move, visibility, protection, or page setup.

## API truths that must be preserved

### Worksheet indexes are zero-based

```csharp
Worksheet firstSheet = workbook.Worksheets[0];
```

Check collection bounds before using a variable index. Prefer name access when the workbook contract guarantees a unique stable sheet name.

### Worksheet names must be unique and valid

Setting `Worksheet.Name` changes the sheet tab name but does not remove its data. Do not assign duplicate names or characters prohibited by spreadsheet formats. Keep names within the supported length.

### Add, copy, and move are different operations

| Goal | Preferred API |
| --- | --- |
| Add an empty named sheet | `workbook.Worksheets.Add("Name")` |
| Copy a sheet in the same workbook | `workbook.Worksheets.AddCopy(indexOrName)` |
| Copy content into an existing sheet | `destination.Copy(source)` |
| Move an existing sheet tab | `worksheet.MoveTo(index)` |

`AddCopy` returns the new worksheet index. Rename the copy deliberately if the generated name is not part of the example.

### Do not remove every visible worksheet

Spreadsheet files generally require at least one visible worksheet. Before hiding or removing sheets, verify that another visible worksheet remains.

### Freeze pane arguments are indexes and pane counts

Use `Worksheet.FreezePanes(row, column, frozenRows, frozenColumns)` deliberately. A common header-row pattern is:

```csharp
worksheet.FreezePanes(1, 0, 1, 0);
```

This freezes one row above zero-based row index `1`. Do not describe index `1` as the first row without explaining the boundary.

### Protection depends on locked cell styles

Cells are locked by default, but locking takes effect only after worksheet protection is enabled. Unlock intended input cells before calling `Worksheet.Protect`.

```csharp
Style style = workbook.CreateStyle();
style.IsLocked = false;
worksheet.Cells["B2"].SetStyle(style);
worksheet.Protect(ProtectionType.All, "change-me", null);
```

Worksheet protection is an editing control, not encryption. Use the encryption category for confidentiality.

### Page setup is worksheet-specific

Headers, footers, margins, orientation, paper size, scaling, print areas, and print titles belong to `Worksheet.PageSetup`. Verify printed or rendered output when page layout is the subject.

### Active, selected, and visible are not synonyms

- `Worksheets.ActiveSheetIndex` identifies the active worksheet.
- `Worksheet.IsSelected` controls tab selection.
- `Worksheet.IsVisible` controls visibility.

Keep at least one sheet active and visible after changes.

## Canonical API map

| API | Purpose | Retrieval aliases |
| --- | --- | --- |
| `Workbook.Worksheets` | Access the workbook's worksheet collection | sheets, tabs, worksheet list |
| `WorksheetCollection.Add` | Add a worksheet | new sheet, create tab |
| `WorksheetCollection.AddCopy` | Duplicate a worksheet | clone sheet, copy tab |
| `WorksheetCollection.RemoveAt` | Remove a worksheet by index | delete sheet |
| `Worksheet` | Configure one worksheet | sheet object, Excel tab |
| `Worksheet.Name` | Read or rename a worksheet | sheet name, tab name |
| `Worksheet.MoveTo` | Reorder a worksheet | move tab, change sheet order |
| `Worksheet.Copy` | Copy source worksheet content and settings | copy between workbooks |
| `Worksheet.FreezePanes` | Keep rows or columns visible while scrolling | freeze header, freeze columns |
| `Worksheet.UnFreezePanes` | Remove frozen panes | unfreeze sheet |
| `Worksheet.Protect` | Apply worksheet editing protection | password-protect sheet |
| `Worksheet.PageSetup` | Configure printed page layout | margins, headers, footer, print area |
| `Cells.InsertRows` / `Cells.DeleteRows` | Insert or remove worksheet rows | shift rows |
| `Cells.InsertColumns` / `Cells.DeleteColumns` | Insert or remove worksheet columns | shift columns |

## Required namespaces

Start with:

```csharp
using System;
using Aspose.Cells;
```

Add `System.IO`, `System.Drawing`, or another framework namespace only when directly used. Add specialized Aspose namespaces only for APIs actually demonstrated.

## Example contract

Every new or regenerated example must:

1. Demonstrate one primary worksheet capability.
2. Be a complete single-file C# program.
3. Use explicit types rather than `var`.
4. Generate workbook data programmatically unless loading is the subject.
5. Use meaningful, deterministic worksheet names and output filenames.
6. Check indexes, names, or collection counts when they come from variables.
7. Preserve at least one visible worksheet.
8. Verify the changed property or collection state.
9. Print a deterministic result or success message.
10. Save a workbook when persistence is relevant.
11. Compile and execute with the repository package and target framework.
12. Match filename, metadata, comments, code, output, and expected result.

## Machine-readable example metadata

New examples should begin with:

```csharp
/*
Title: Add and rename an Excel worksheet in C#
Intent: Add a worksheet, populate it, and verify its name
Category: working-with-worksheets
Primary API: WorksheetCollection.Add
Secondary APIs: Worksheet.Name, Cell.PutValue, Workbook.Save
Input: Programmatically generated workbook
Output: worksheet-result.xlsx
Expected Result: Workbook contains a worksheet named Sales
Product: Aspose.Cells for .NET
Language: C#
*/
```

Metadata must describe the code exactly and use canonical API casing.

## Patterns by task

### Access by name with a null check

```csharp
Worksheet worksheet = workbook.Worksheets["Sales"];
if (worksheet == null)
{
    throw new InvalidOperationException("Worksheet 'Sales' was not found.");
}
```

### Copy within a workbook

```csharp
int copyIndex = workbook.Worksheets.AddCopy("Sales");
Worksheet copy = workbook.Worksheets[copyIndex];
copy.Name = "Sales Copy";
```

### Move a worksheet

```csharp
Worksheet worksheet = workbook.Worksheets["Sales"];
worksheet.MoveTo(0);
```

### Hide a worksheet safely

```csharp
if (workbook.Worksheets.Count > 1)
{
    workbook.Worksheets["Lookup"].IsVisible = false;
}
```

### Configure page layout

```csharp
PageSetup pageSetup = worksheet.PageSetup;
pageSetup.Orientation = PageOrientationType.Landscape;
pageSetup.FitToPagesWide = 1;
pageSetup.FitToPagesTall = 0;
```

## Verification requirements

Verify state in memory before save and reload the saved workbook when persistence is central.

Useful checks include:

- Worksheet count, name, index, and order
- Cell values or formulas after a copy
- `IsVisible`, `IsSelected`, or active-sheet index
- Frozen-pane state using the package-supported query API
- Protection and unlocked-cell styles
- Page setup properties and rendered page count
- Output file existence and nonzero size

Do not treat "no exception" as sufficient validation.

## Performance and reliability

- Cache worksheet references instead of repeatedly resolving names in large loops.
- Avoid `AutoFitRows` or `AutoFitColumns` across large unused ranges.
- Batch row, column, style, and page-setup changes where possible.
- Copy only needed worksheets when consolidating large workbooks.
- Sanitize user-supplied sheet names and resolve collisions deterministically.
- Do not share a mutable `Workbook` across threads.
- Measure memory and elapsed time for large-sheet operations.

## Security and compliance

- Never hard-code production passwords or secrets in examples.
- Explain that worksheet protection does not encrypt file contents.
- Validate output paths and avoid overwriting source workbooks unintentionally.
- Treat external hyperlinks, formulas, macros, and embedded objects in loaded sheets as untrusted input.
- Remove sensitive hidden worksheets only when the task explicitly requires it; hiding is not data redaction.

## Anti-patterns

Do not:

- Use Excel Interop or require Microsoft Excel.
- Use an index without validating it when the index is dynamic.
- Hide or remove the last visible worksheet.
- Claim that renaming or moving a worksheet changes its cell data.
- Call worksheet protection encryption.
- Set a cell's locked style after protection and assume earlier behavior was validated.
- Add charts, tables, images, formulas, and unrelated features to a simple worksheet example.
- Depend on `input.xlsx` unless loading an existing workbook is the purpose.
- Swallow exceptions and print success.
- Invent properties from task wording.

## Review checklist

- [ ] One worksheet operation is dominant.
- [ ] APIs exist in the installed Aspose.Cells version.
- [ ] Index and name assumptions are safe.
- [ ] At least one visible worksheet remains.
- [ ] Verification checks the claimed state.
- [ ] Output filename is deterministic.
- [ ] Code uses explicit types and minimal namespaces.
- [ ] The example compiles and runs.
- [ ] Comments and metadata match behavior.
- [ ] No secrets or unrelated dependencies are introduced.

## Retrieval guidance for AI systems

Prefer answers in this order:

1. Match the exact operation: add, access, rename, copy, move, hide, freeze, protect, or page setup.
2. Select the smallest API that performs that operation.
3. Preserve sheet visibility, naming, and index invariants.
4. Include an observable verification.
5. Save only when a persisted workbook is part of the result.

Useful aliases:

- create Excel sheet in C#
- duplicate worksheet with Aspose.Cells
- reorder Excel sheet tabs
- freeze Excel header row without Interop
- hide or unhide worksheet
- protect worksheet and unlock input cells
- Excel page setup in .NET

## Related categories

- [`manage-workbook`](../manage-workbook/) - workbook-level structure and settings
- [`rows-and-columns`](../rows-and-columns/) - row and column operations
- [`cells-data`](../cells-data/) - read, write, import, and export cells
- [`format-cells`](../format-cells/) - cell and range styling
- [`working-with-tables`](../working-with-tables/) - structured Excel tables
- [`working-with-charts`](../working-with-charts/) - worksheet charts
- [`working-with-shapes`](../working-with-shapes/) - worksheet drawing objects

## Official Aspose.Cells resources

- [Manage Worksheets documentation](https://docs.aspose.com/cells/net/managing-worksheets/)
- [Worksheet API](https://reference.aspose.com/cells/net/aspose.cells/worksheet/)
- [WorksheetCollection API](https://reference.aspose.com/cells/net/aspose.cells/worksheetcollection/)
- [PageSetup API](https://reference.aspose.com/cells/net/aspose.cells/pagesetup/)
- [Worksheet protection documentation](https://docs.aspose.com/cells/net/protecting-worksheets/)
- [Aspose.Cells NuGet package](https://www.nuget.org/packages/Aspose.Cells/)

## Final authority

The installed Aspose.Cells package and official API reference are authoritative. Existing generated files may contain speculative, version-sensitive, third-party, or externally dependent scenarios. Validate API signatures, compile, execute, and inspect the claimed result before accepting or featuring any example.

