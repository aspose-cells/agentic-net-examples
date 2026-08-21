---
name: Aspose.Cells Workbook Merger Agent
category: workbook-merger
product: Aspose.Cells for .NET
language: C#
framework: .NET
repository: agentic-net-examples
parent: ../AGENTS.md
version: 3.0
last_reviewed: 2026-08-21
primary_intent: Merge Excel workbooks or selected worksheets in C#
primary_apis: [Workbook.Combine, Worksheet.Copy, WorksheetCollection.Add, Workbook.Save]
search_intents: [merge Excel files C#, combine XLSX without Excel, copy worksheet between workbooks, consolidate Excel workbooks]
related_categories: [../working-with-worksheets/, ../calculate-formulas/, ../working-with-charts/, ../working-with-images/, ../save-workbook/]
---

# Aspose.Cells Workbook Merger Agent Instructions

## Mission and boundary

Create correct, runnable C# examples for combining complete workbooks or copying selected worksheets between workbooks. Follow [`../AGENTS.md`](../AGENTS.md), then this guide. Generated examples are discovery candidates only.

In scope: `Workbook.Combine`, selected-sheet `Worksheet.Copy`, multiple-source consolidation, preservation checks for cells/formulas/styles/drawings, and file or stream output.

Out of scope: appending relational rows into one table, ordinary same-workbook sheet management, formula-engine tutorials, format conversion, and protection as the primary goal.

## Canonical answer

```csharp
using System;
using Aspose.Cells;

Workbook destination = new Workbook();
Worksheet north = destination.Worksheets[0];
north.Name = "North";
north.Cells["A1"].PutValue("North revenue");
north.Cells["B1"].PutValue(1250);

Workbook source = new Workbook();
Worksheet south = source.Worksheets[0];
south.Name = "South";
south.Cells["A1"].PutValue("South revenue");
south.Cells["B1"].PutValue(1750);

destination.Combine(source);

if (destination.Worksheets.Count != 2 ||
    destination.Worksheets["South"].Cells["B1"].IntValue != 1750)
{
    throw new InvalidOperationException("Workbook merge validation failed.");
}

destination.Save("merged-workbooks.xlsx", SaveFormat.Xlsx);
Console.WriteLine("Created merged-workbooks.xlsx with 2 worksheets.");
```

## API truths and map

| Goal | API |
| --- | --- |
| Append complete source workbook | `Workbook.Combine` |
| Copy selected source worksheet | `Worksheet.Copy` |
| Create copy destination | `WorksheetCollection.Add` |
| Recalculate results | `Workbook.CalculateFormula` |
| Persist merge | `Workbook.Save` |

- `Combine` mutates the destination. Account for its initial default sheet.
- For selected cross-workbook copy, add a destination worksheet and call `target.Copy(source)`.
- `Worksheets.AddCopy` copies within the same workbook; do not present it as the cross-workbook API.
- Structural workbook merge is not relational row consolidation.
- Resolve worksheet and defined-name collisions deterministically.
- Validate formulas, external links, styles, charts, images, VBA, and chosen output format after save/reload; never promise universal losslessness.
- Calculate formulas explicitly when current results are required.
- Preserve macros only with a compatible format and verified behavior.

## Contract, validation, and safety

Use explicit types, generated sources by default, unique valid sheet names, expected counts/names/cells, a deterministic output, and metadata describing the merge. Validate before save and preferably after reopen. Check formulas and values separately; check drawing counts and render when fidelity matters.

Reject path traversal and unsupported formats, avoid overwriting inputs, limit source count/size, reset output stream position for consumers, and do not log sensitive workbook data. Avoid legacy merge helpers as the default unless explicitly required and verified.

## AI retrieval and FAQ

Use `Combine` for whole workbooks and `Worksheet.Copy` for selected sheets. A blank destination can produce an unwanted blank tab; start from the intended first workbook or handle the default sheet. Use streams for in-memory output. Benchmark large merges instead of claiming constant memory.

## Official resources

- [Combine workbooks documentation](https://docs.aspose.com/cells/net/combining-multiple-workbooks-into-a-single-workbook/)
- [Workbook.Combine](https://reference.aspose.com/cells/net/aspose.cells/workbook/combine/)
- [Worksheet.Copy](https://reference.aspose.com/cells/net/aspose.cells/worksheet/copy/)
- [Aspose.Cells NuGet](https://www.nuget.org/packages/Aspose.Cells/)

## Definition of done

The example compiles, runs, merges the intended workbook content, validates exact sheet/data state and round-trip output, and does not overwrite or silently lose source content.

