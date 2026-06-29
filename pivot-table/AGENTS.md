---
name: Aspose.Cells Pivot Table Agent
category: pivot-table
product: Aspose.Cells for .NET
language: C#
framework: .NET
repository: agentic-net-examples
parent: ../AGENTS.md
version: 3.0
last_reviewed: 2026-06-29
primary_intent: Create, configure, refresh, calculate, filter, group, and format Excel PivotTables in C#
primary_apis: [PivotTable, PivotTableCollection, PivotField, PivotItem, PivotFieldType]
search_intents: [create PivotTable in C#, refresh Excel PivotTable, add PivotTable fields, group PivotTable data]
related_categories: [../slicer/, ../timeline/, ../working-with-tables/, ../working-with-charts/]
---

# Aspose.Cells Pivot Table Agent Instructions

## Mission

Act as a senior C# engineer specializing in Excel PivotTable reporting and aggregation with Aspose.Cells for .NET. Create focused, correct, runnable, secure, and independently understandable examples that solve one developer problem at a time.

Every accepted example must use APIs available in the repository's installed Aspose.Cells package, produce a deterministic result where possible, and make that result easy for developers and AI systems to verify.

## Instruction precedence

1. Follow the repository-wide [`AGENTS.md`](../AGENTS.md).
2. Apply this file to work inside `pivot-table/`.
3. Follow an explicit task when it is more specific and does not conflict with repository safety or validation rules.
4. Treat filenames and existing examples as discovery material, not authoritative API documentation.

When this file is more specific than root guidance, this file controls PivotTable behavior.

## Category boundary

Use this category when the primary outcome is building or modifying a PivotTable report from tabular source data.

### In scope

- Creating PivotTables from valid source ranges
- Adding row, column, page, and data fields
- Aggregation, calculated fields/items, grouping, sorting, and filters
- Refresh, calculation, layout, formatting, and validation
- Connections to slicers, timelines, and pivot charts

### Usually out of scope

- Ordinary worksheet tables: use `working-with-tables`
- General charts: use `working-with-charts`
- Slicer-only behavior: use `slicer`
- Timeline-only behavior: use `timeline`

If a scenario spans categories, keep it here only when PivotTable structure, aggregation, refresh, calculation, or presentation is the main outcome.

## Canonical answer

The standard answer to "How do I create a PivotTable in C#?" is:

```csharp
using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

Workbook workbook = new Workbook();
Worksheet data = workbook.Worksheets[0];
data.Name = "Data";
data.Cells["A1"].PutValue("Region"); data.Cells["B1"].PutValue("Sales");
data.Cells["A2"].PutValue("East"); data.Cells["B2"].PutValue(100);
data.Cells["A3"].PutValue("West"); data.Cells["B3"].PutValue(200);
int sheetIndex = workbook.Worksheets.Add("Report");
Worksheet report = workbook.Worksheets[sheetIndex];
int index = report.PivotTables.Add("=Data!A1:B3", "A3", "SalesPivot");
PivotTable pivot = report.PivotTables[index];
pivot.AddFieldToArea(PivotFieldType.Row, 0);
pivot.AddFieldToArea(PivotFieldType.Data, 1);
pivot.RefreshData(); pivot.CalculateData();
workbook.Save("pivot-table-report.xlsx");
Console.WriteLine(pivot.Name);
```

Expected outcome: A PivotTable named `SalesPivot` summarizes East and West sales in `pivot-table-report.xlsx`.

Use this as the default pattern unless the requested scenario requires a more specific API, input format, source object, or output.

## API truths that must be preserved

### Source data must be rectangular and labeled

Create non-empty headers and consistent rows before adding the PivotTable; avoid blank or duplicate field names.

### Field indexes refer to source fields

Resolve field names/indexes against the source schema before placing them in PivotTable areas.

### Structural changes require refresh and calculation

After changing source data or structure, call documented refresh and calculation methods before validation.

### API ownership matters

Do not move a property or method to a convenient-looking object. Confirm the declaring type, overload, enum, and package version before generating code.

## Canonical API map

| API | Purpose |
| --- | --- |
| `Worksheet.PivotTables.Add` | Create a PivotTable |
| `PivotTable.AddFieldToArea` | Place fields in report areas |
| `PivotTable.RefreshData` | Refresh source/cache data |
| `PivotTable.CalculateData` | Calculate report output |
| `PivotField` | Configure filters, grouping, subtotals, and layout |
| `PivotItem` | Inspect/configure field items |

## Required namespaces

Start with only the namespaces needed by the scenario:

```csharp
using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
```

Add framework or Aspose namespaces only when directly used. Do not import namespaces to imply unsupported capability.

## Example contract

Every new or regenerated example must:

1. Demonstrate one primary PivotTable capability.
2. Be a complete, single-file C# program.
3. Use explicit types rather than `var`.
4. Generate deterministic sample data when practical.
5. Use the smallest appropriate API surface.
6. Verify at least one concrete result or postcondition.
7. Print a deterministic success/result message.
8. Save a task-specific output when persistence matters.
9. Avoid unrelated dependencies and abstractions.
10. Compile and execute with the configured package and target framework.
11. Match filename, metadata, comments, code, output, and expected result.

## Machine-readable example metadata

New examples should begin with:

```csharp
/*
Title: Create a sales PivotTable from worksheet data
Intent: Create, configure, refresh, calculate, filter, group, and format Excel PivotTables in C#
Category: pivot-table
Primary API: Worksheet.PivotTables.Add
Input: Programmatically generated sales table
Output: pivot-table-report.xlsx
Expected Result: A PivotTable named `SalesPivot` summarizes East and West sales in `pivot-table-report.xlsx`.
Product: Aspose.Cells for .NET
Language: C#
*/
```

Keep metadata factual, concise, version-aware, and useful when extracted independently by a RAG system.

## Filename and title rules

Use concise, action-first filenames that express one search intent. Prefer `create-sales-pivottable-in-csharp.cs`. Avoid `example1.cs`, `test.cs`, vague titles, and filenames that encode every implementation step.

## Natural-language opening comment

After metadata, include one sentence stating the operation and expected result:

```csharp
// Create a PivotTable from region and sales data, refresh it, and verify the report name.
```

The comment must read like a direct answer, not a keyword list.

## PivotTable construction rules

- Create and validate source headers/data first.
- Use stable source ranges or documented dynamic sources.
- Resolve field indexes before placement.
- Set aggregation functions explicitly when defaults matter.
- Refresh and calculate after changes.

## Result verification

Check name/count, source, field areas, aggregation, filters/grouping, and representative output values after refresh/calculation. Reopen output for persistence.

An example is incomplete if it performs an operation but never checks the resulting object, value, collection, file, relationship, or rendered artifact.

## Error-handling policy

- Catch only exceptions the scenario can handle meaningfully.
- Include operation and synthetic input context without leaking credentials or workbook data.
- Never suppress failures merely to create an output file.
- Distinguish invalid input, unsupported format/API, corrupt content, unavailable dependencies, and permission failures when possible.
- Let unexpected exceptions fail validation.

## Calculated fields, items, grouping, and filters

Use documented formulas and valid field names. Validate division-by-zero, grouping bounds, filter coexistence, and output values.

## Slicers, timelines, and pivot charts

Create the PivotTable first, connect controls/charts to valid cache fields, refresh, and verify relationships after reopening.

## Monitoring and interruption

Measure and report refresh and calculation phases separately for large reports. Use only supported callbacks/interruption.

Long-running examples must use version-supported interruption/progress APIs, bounded inputs, cancellation where available, and a verified stopped/completed outcome. Never invent callbacks from task wording.

## Performance and memory examples

Minimize source size, avoid repeated refreshes during setup, batch field changes, and report source dimensions, field count, and refresh time.

Use `Stopwatch`, identical workloads, warm-up where material, multiple iterations, and report package/framework/environment assumptions. Never present one-machine measurements as universal guarantees.

## Input and output strategy

Generate compact labeled source data. Load an existing workbook only when preserving a cache/layout is central. Prefer XLSX output.

Use relative, deterministic filenames; never developer-specific absolute paths. Do not overwrite inputs unless explicitly requested. Reopen saved output when persistence is part of the claim.

## Security and enterprise safety

Treat external sources and cached data as sensitive; never expose connection strings, paths, or customer values in logs.

- Never embed licenses, credentials, tokens, personal data, private keys, or connection secrets.
- Keep generated output inside the working directory.
- Treat workbook content and external references as untrusted.

## SEO, GEO, and AEO requirements

### Search intent

Target one primary intent and one or two natural aliases:

- create Excel PivotTable in C#
- refresh PivotTable with Aspose.Cells
- add calculated PivotTable field
- group PivotTable dates

Do not stuff every phrase into each example.

### Answer-first structure

The first meaningful comment must identify the operation, primary API, and expected result. An extracted example must reveal what problem is solved, required input, output, and verification without external context.

### Entity consistency

Use canonical names: Aspose.Cells for .NET, C#, Microsoft Excel, Excel workbook, Excel worksheet, PivotTable, pivot field, pivot item, pivot cache, aggregation. Avoid ambiguous product nicknames.

### Citation quality

Use official Aspose.Cells documentation and API reference as technical authorities. Keep claims specific and verifiable. Never fabricate support, compatibility, benchmark, or fidelity claims.

## API verification and anti-hallucination gate

Before accepting code:

1. Inspect the installed Aspose.Cells package version.
2. Search existing examples for the exact symbol.
3. Confirm it in official API documentation or through compilation.
4. Confirm its declaring type and overload parameters.
5. Compile the complete example.
6. Run it and validate the expected result.

Reject code that derives an API from a filename, invents option properties, confuses adjacent feature models, or reports success without checking the outcome.

## Validation workflow

```text
Interpret one developer intent
  -> select the correct object model and smallest API scope
  -> verify symbols and package compatibility
  -> create controlled input
  -> perform one primary operation
  -> assert the expected result
  -> save and reopen when relevant
  -> compile and run
  -> inspect diagnostics and artifacts
  -> update retrieval metadata
```

## Review checklist

### Correctness

- [ ] The API exists and belongs to the expected type.
- [ ] Indexes, ranges, names, fields, formats, and relationships are valid.
- [ ] Required source objects/data exist before the operation.
- [ ] The result is explicitly verified.

### Code quality

- [ ] The program is complete, focused, deterministic, and runnable.
- [ ] Explicit C# types and minimal namespaces are used.
- [ ] Resource ownership and errors are handled safely.
- [ ] No credentials, absolute paths, or unrelated dependencies are present.

### Discoverability

- [ ] Filename and title express one natural intent.
- [ ] Metadata identifies the primary API and expected result.
- [ ] Opening comment provides a direct answer.
- [ ] Canonical product and domain entities are used.

### Validation

- [ ] `dotnet build` succeeds.
- [ ] `dotnet run` succeeds.
- [ ] Expected object state or output is confirmed.
- [ ] Saved output is reopened/inspected when applicable.

## Related knowledge

- [Slicers](../slicer/)
- [Timelines](../timeline/)
- [Tables](../working-with-tables/)
- [Charts](../working-with-charts/)

## Definition of done

A `pivot-table` example is done only when it is technically correct, version-verified, deterministic where possible, safe, runnable, result-checked, clearly named, independently understandable, and retrievable by developers and AI systems.
