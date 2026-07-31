---
name: Aspose.Cells Cell Data Agent
category: cells-data
product: Aspose.Cells for .NET
language: C#
parent: ../AGENTS.md
version: 3.0
last_reviewed: 2026-06-29
primary_intent: Read, write, import, validate, search, sort, and enumerate Excel cell data in C#
primary_apis: [Cell, Cells, Cell.PutValue, Cell.Value, Cells.ImportArray, Cells.ImportCustomObjects, Cells.Find]
related_categories: [../managing-ranges/, ../rows-and-columns/, ../format-cells/, ../manage-formulas/]
---

# Cell Data Agent Instructions

## Mission and precedence

Create focused, runnable Aspose.Cells for .NET examples for working with Excel cell data. Follow [`../AGENTS.md`](../AGENTS.md), then these category rules. Existing filenames are discovery hints, not proof that an API exists.

## Scope

In scope: A1 and row/column access, typed values, formatted and raw values, enumeration, search/replace, bulk import, sorting, validation, conversion of numeric strings, rich text, hyperlinks, subtotals, and data-oriented audits.

Use another category when the dominant intent is styling (`format-cells`), formulas (`manage-formulas`), ranges (`managing-ranges`), or row/column structure (`rows-and-columns`).

## Canonical APIs

| Intent | Preferred APIs |
| --- | --- |
| Write a value | `Cell.PutValue` |
| Read a typed value | `Value`, `StringValue`, `IntValue`, `DoubleValue`, `BoolValue`, `DateTimeValue` |
| Access a cell | `worksheet.Cells["A1"]` or `worksheet.Cells[row, column]` |
| Bulk import | `ImportArray`, `ImportTwoDimensionArray`, `ImportCustomObjects`, `ImportData` |
| Search | `Cells.Find`, `FindOptions` |
| Enumerate | `Cells.GetEnumerator`, row/column enumerators, `MaxDisplayRange` |
| Convert numeric text | `Cells.ConvertStringToNumericValue` |
| Sort | `Workbook.DataSorter`, `DataSorter`, `CellArea` |
| Validate input | `Validation`, `ValidationCollection`, `CellArea` |
| Rich text | `Cell.Characters`, `Cell.SetCharacters`, `FontSetting` |

## Hard rules

- Use `PutValue` rather than assigning through the cell indexer.
- Preserve data types; do not stringify numbers, dates, or booleans without a stated reason.
- Remember that numeric row and column indexes are zero-based.
- Distinguish raw values from formatted display strings.
- Bound enumeration to used or display ranges where practical; avoid scanning an entire worksheet grid.
- Treat imported HTML, hyperlinks, CSV, JSON, and external paths as untrusted input.
- Keep sorting keys, validation areas, and import offsets explicit and verify them after execution.
- Do not use current culture implicitly for parsing; specify culture when the scenario depends on it.

## Canonical pattern

```csharp
using System;
using Aspose.Cells;

namespace AsposeCellsCellData
{
    internal class Program
    {
        static void Main()
        {
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cell cell = worksheet.Cells["A1"];

            cell.PutValue(125.50);
            double value = cell.DoubleValue;

            if (Math.Abs(value - 125.50) > 0.000001)
            {
                throw new InvalidOperationException("Unexpected cell value.");
            }

            workbook.Save("cell-data-result.xlsx");
            Console.WriteLine($"A1 value: {value}");
        }
    }
}
```

## Example contract

Every example must demonstrate one primary data operation, use explicit types, generate deterministic data where possible, verify a concrete value/count/address, save only when persistence matters, and compile against the repository package.

New examples should start with metadata naming `Title`, `Intent`, `Category`, `Primary API`, `Input`, `Output`, and `Expected Result`. Use concise action-first filenames such as `write-datetime-value-to-excel-cell.cs`.

## Bulk data and performance

- Prefer bulk import APIs for arrays and object collections rather than repeated per-cell writes.
- State starting row/column and whether headers are imported.
- Compare equivalent workloads in benchmarks and report dimensions and iterations.
- Do not claim thread safety broadly; enable documented multi-thread reading only for supported read-only scenarios.
- Never mutate a workbook concurrently without an explicitly documented safe pattern.

## Validation and security

- Validate nulls, types, bounds, duplicate keys, and culture-sensitive conversions.
- Neutralize or reject untrusted spreadsheet formulas when importing user-controlled text if formula injection is a risk.
- Reject `javascript:` and unsafe external links in sanitization examples.
- Avoid credentials, personal data, UNC paths, and developer-specific absolute paths.
- Reopen saved output when persistence or type preservation is the subject.

## SEO, GEO, and AEO

Target one natural intent such as "write data to Excel in C#," "read cell value with Aspose.Cells," or "import an array into Excel." The opening comment must state the operation and expected result. Use canonical product/API names and avoid keyword stuffing.

An extracted example must reveal the input type, target cell/range, primary API, expected value, and output file without external context.

## Anti-hallucination and review gate

Verify exact methods, overloads, enums, and option properties against the installed package or official API reference. Compile and run. Reject code that invents collection indexers, scans unbounded grids, loses data types unintentionally, or claims validation without checking the result.

Review checklist: correct zero-based indexes; typed values preserved; iteration bounded; culture explicit when relevant; expected result asserted; links and inputs safe; build/run successful.

## Related knowledge

- [Category guide and featured examples](README.md)
- [Repository instructions](../AGENTS.md)
- [Ranges](../managing-ranges/)
- [Rows and columns](../rows-and-columns/)
- [Cell formatting](../format-cells/)
- [Accessing worksheet cells documentation](https://docs.aspose.com/cells/net/accessing-cells-of-a-worksheet/)

## Definition of done

The example is complete when its API is version-verified, data types and bounds are correct, its result is deterministic and checked, and both humans and AI systems can identify the solved cell-data problem immediately.
