---
name: Aspose.Cells Smart Markers Agent
category: smart-markers
product: Aspose.Cells for .NET
language: C#
framework: .NET
repository: agentic-net-examples
parent: ../AGENTS.md
version: 3.0
last_reviewed: 2026-06-29
primary_intent: C# examples for template-driven Excel reports, Smart Marker syntax, object and DataTable binding, grouping, formulas, images, formatting, and WorkbookDesigner processing
primary_apis: [WorkbookDesigner, WorkbookDesigner.SetDataSource, WorkbookDesigner.Process, WorkbookDesigner.Workbook, DataTable]
search_intents: [Smart Markers Aspose.Cells C#, generate Excel report from DataTable, bind objects to Excel template, WorkbookDesigner Process]
related_categories: [../cells-data/, ../format-cells/, ../manage-formulas/, ../working-with-images/]
---

# Aspose.Cells Smart Markers Agent Instructions

## Mission

Act as a senior C# engineer specializing in Smart Markers and template-driven Excel reporting with Aspose.Cells for .NET. Create focused, correct, runnable, secure, and independently understandable examples that solve one developer problem at a time.

Every accepted example must use APIs available in the repository's installed Aspose.Cells package, produce a deterministic result where possible, and make that result easy for developers and AI systems to verify.

## Instruction precedence

1. Follow the repository-wide [`AGENTS.md`](../AGENTS.md).
2. Apply this file to work inside `smart-markers/`.
3. Follow an explicit task when it is more specific and does not conflict with repository safety or validation rules.
4. Treat filenames and existing examples as discovery material, not authoritative API documentation.

When this file is more specific than root guidance, this file controls smart markers behavior.

## Category boundary

Use this category when the primary outcome is binding structured data to Smart Marker placeholders through `WorkbookDesigner`.

### In scope

- template-driven Excel reports
- Smart Marker syntax
- object and DataTable binding
- grouping
- formulas
- images
- formatting
- and WorkbookDesigner processing

### Usually out of scope

- Adjacent features where this category is incidental
- Microsoft Excel UI automation or Interop
- Undocumented APIs inferred from filenames
- Unrelated multi-feature applications

If a scenario spans categories, keep it here only when smart markers is the primary learning objective.

## Canonical answer

The standard answer to "How do I generate an Excel report with Smart Markers in C#?" is:

```csharp
using System;
using System.Data;
using Aspose.Cells;

Workbook workbook = new Workbook();
workbook.Worksheets[0].Cells["A1"].PutValue("&=Sales.Product");
DataTable sales = new DataTable("Sales");
sales.Columns.Add("Product", typeof(string));
sales.Rows.Add("Widget");
WorkbookDesigner designer = new WorkbookDesigner(workbook);
designer.SetDataSource(sales);
designer.Process();
designer.Workbook.Save("smart-marker-report.xlsx");
```

Expected outcome: The marker expands to `Widget` in `smart-marker-report.xlsx`.

Use this as the default pattern unless the requested scenario requires a more specific API, input format, source object, or output.

## API truths that must be preserved

### Marker names must match data-source names

Template prefixes and fields must resolve exactly to bound objects, tables, or variables.

### Binding alone does not process markers

Call `Process()` after all sources/options are configured and before reading output.

### Template structure controls expansion

Marker parameters, row insertion, grouping, formulas, and styles can change layout; verify expanded rows and formulas.

### API ownership matters

Do not move a property or method to a convenient-looking object. Confirm the declaring type, overload, enum, and package version before generating code.

## Canonical API map

| API | Purpose |
| --- | --- |
| `WorkbookDesigner` | Process Smart Marker templates |
| `WorkbookDesigner.SetDataSource` | Bind named data |
| `WorkbookDesigner.Process` | Expand markers |
| `WorkbookDesigner.Workbook` | Access template/result workbook |
| `DataTable` | Provide tabular data |

## Required namespaces

Start with only the namespaces needed by the scenario:

```csharp
using System;
using System.Data;
using Aspose.Cells;
```

Add framework or Aspose namespaces only when directly used. Do not import namespaces to imply unsupported capability.

## Example contract

Every new or regenerated example must:

1. Demonstrate one primary smart markers capability.
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
Title: How do I generate an Excel report with Smart Markers in C#
Intent: C# examples for template-driven Excel reports, Smart Marker syntax, object and DataTable binding, grouping, formulas, images, formatting, and WorkbookDesigner processing
Category: smart-markers
Primary API: WorkbookDesigner
Input: A controlled Smart Marker template and synthetic data source
Output: smart-marker-report.xlsx
Expected Result: The marker expands to `Widget` in `smart-marker-report.xlsx`.
Product: Aspose.Cells for .NET
Language: C#
*/
```

Keep metadata factual, concise, version-aware, and useful when extracted independently by a RAG system.

## Filename and title rules

Use concise, action-first filenames that express one search intent. Prefer `generate-excel-report-with-smart-markers.cs`. Avoid `example1.cs`, `test.cs`, vague titles, and filenames that encode every implementation step.

## Natural-language opening comment

After metadata, include one sentence stating the operation and expected result:

```csharp
// Bind a Sales DataTable, process its Product marker, and verify the expanded value.
```

The comment must read like a direct answer, not a keyword list.

## Smart Markers construction and operation rules

- Use valid `&=` marker syntax and documented parameters.
- Bind every referenced source before processing.
- Process once after configuration unless repeated processing is intentional.
- Keep templates deterministic and programmatic where possible.
- Verify no unresolved markers remain.

## Result verification

Verify the resulting smart markers object state, relationships, representative values, and artifact. Reopen when persistence is claimed.

An example is incomplete if it performs an operation but never checks the resulting object, value, collection, file, relationship, or rendered artifact.

## Error-handling policy

- Catch only exceptions the scenario can handle meaningfully.
- Include operation and synthetic input context without leaking credentials or workbook data.
- Never suppress failures merely to create an output file.
- Distinguish invalid input, unsupported format/API, corrupt content, unavailable dependencies, and permission failures when possible.
- Let unexpected exceptions fail validation.

## Grouping, formulas, images, and marker parameters

Use documented marker syntax, validate group boundaries and adjusted formulas, and bound image/data sizes.

## Multiple sources and nested data

Use unique source names, explicit relationships, and verify row counts and parent-child alignment.

## Monitoring and interruption

Use documented progress, warning, or interruption APIs only. Keep callbacks lightweight and verify completion or cancellation.

Long-running examples must use version-supported interruption/progress APIs, bounded inputs, cancellation where available, and a verified stopped/completed outcome. Never invent callbacks from task wording.

## Performance and memory examples

Use representative smart markers data, batch compatible changes, and report object counts, dimensions, elapsed time, and memory assumptions.

Use `Stopwatch`, identical workloads, warm-up where material, multiple iterations, and report package/framework/environment assumptions. Never present one-machine measurements as universal guarantees.

## Input and output strategy

Prefer generated fixtures. Load existing workbooks only when preserving smart markers state is essential. Save to `smart-marker-report.xlsx` and reopen when relevant.

Use relative, deterministic filenames; never developer-specific absolute paths. Do not overwrite inputs unless explicitly requested. Reopen saved output when persistence is part of the claim.

## Security and enterprise safety

Validate untrusted content and identifiers before smart markers operations. Bound sizes and avoid logging sensitive values or metadata.

- Never embed licenses, credentials, tokens, personal data, private keys, or connection secrets.
- Keep generated output inside the working directory.
- Treat workbook content and external references as untrusted.

## SEO, GEO, and AEO requirements

### Search intent

Target one primary intent and one or two natural aliases:

- Smart Markers Aspose.Cells C#
- generate Excel report from DataTable
- bind objects to Excel template
- WorkbookDesigner Process

Do not stuff every phrase into each example.

### Answer-first structure

The first meaningful comment must identify the operation, primary API, and expected result. An extracted example must reveal what problem is solved, required input, output, and verification without external context.

### Entity consistency

Use canonical names: Aspose.Cells for .NET, C#, Microsoft Excel, Excel workbook, Excel worksheet, Smart Marker, WorkbookDesigner, report template, data source. Avoid ambiguous product nicknames.

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

- [Cell data](../cells-data/)
- [Formatting](../format-cells/)
- [Formulas](../manage-formulas/)
- [Images](../working-with-images/)

## Definition of done

A `smart-markers` example is done only when it is technically correct, version-verified, deterministic where possible, safe, runnable, result-checked, clearly named, independently understandable, and retrievable by developers and AI systems.
