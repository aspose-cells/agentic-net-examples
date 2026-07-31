---
name: Aspose.Cells Timelines Agent
category: timeline
product: Aspose.Cells for .NET
language: C#
framework: .NET
repository: agentic-net-examples
parent: ../AGENTS.md
version: 3.0
last_reviewed: 2026-06-29
primary_intent: C# examples for creating, accessing, positioning, styling, connecting, removing, and rendering Excel PivotTable timelines based on date fields
primary_apis: [Worksheet.Timelines, TimelineCollection.Add, Timeline, PivotTable]
search_intents: [create Excel timeline in C#, add timeline to PivotTable, filter PivotTable by date, remove Excel timeline]
related_categories: [../pivot-table/, ../slicer/, ../working-with-pdf/, ../working-with-charts/]
---

# Aspose.Cells Timelines Agent Instructions

## Mission

Act as a senior C# engineer specializing in Excel PivotTable timeline controls with Aspose.Cells for .NET. Create focused, correct, runnable, secure, and independently understandable examples that solve one developer problem at a time.

Every accepted example must use APIs available in the repository's installed Aspose.Cells package, produce a deterministic result where possible, and make that result easy for developers and AI systems to verify.

## Instruction precedence

1. Follow the repository-wide [`AGENTS.md`](../AGENTS.md).
2. Apply this file to work inside `timeline/`.
3. Follow an explicit task when it is more specific and does not conflict with repository safety or validation rules.
4. Treat filenames and existing examples as discovery material, not authoritative API documentation.

When this file is more specific than root guidance, this file controls timelines behavior.

## Category boundary

Use this category when the primary outcome is creating or managing an Excel timeline connected to a PivotTable date field.

### In scope

- creating
- accessing
- positioning
- styling
- connecting
- removing
- and rendering Excel PivotTable timelines based on date fields

### Usually out of scope

- Adjacent features where this category is incidental
- Microsoft Excel UI automation or Interop
- Undocumented APIs inferred from filenames
- Unrelated multi-feature applications

If a scenario spans categories, keep it here only when timelines is the primary learning objective.

## Canonical answer

The standard answer to "How do I add a timeline to a PivotTable in C#?" is:

```csharp
using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Pivot;

Workbook workbook = new Workbook("pivot-with-dates.xlsx");
Worksheet worksheet = workbook.Worksheets["Report"];
PivotTable pivot = worksheet.PivotTables[0];
int index = worksheet.Timelines.Add(pivot, "H2", "Date");
Timeline timeline = worksheet.Timelines[index];
workbook.Save("pivot-timeline.xlsx");
Console.WriteLine(timeline.Name);
```

Expected outcome: A timeline linked to the PivotTable date field is saved at H2.

Use this as the default pattern unless the requested scenario requires a more specific API, input format, source object, or output.

## API truths that must be preserved

### A timeline requires a PivotTable date field

Create/load the PivotTable and confirm the selected source field contains valid dates.

### A timeline is an Excel filter control

It filters supported PivotTable cache data; it is not a general-purpose chart or project timeline.

### Rendered timelines are static

PDF/image output can display appearance but does not preserve Excel interaction.

### API ownership matters

Do not move a property or method to a convenient-looking object. Confirm the declaring type, overload, enum, and package version before generating code.

## Canonical API map

| API | Purpose |
| --- | --- |
| `Worksheet.Timelines` | Access timeline controls |
| `TimelineCollection.Add` | Create a timeline |
| `Timeline` | Configure timeline properties |
| `PivotTable` | Provide source cache and date field |

## Required namespaces

Start with only the namespaces needed by the scenario:

```csharp
using System;
using Aspose.Cells;
```

Add framework or Aspose namespaces only when directly used. Do not import namespaces to imply unsupported capability.

## Example contract

Every new or regenerated example must:

1. Demonstrate one primary timelines capability.
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
Title: How do I add a timeline to a PivotTable in C#
Intent: C# examples for creating, accessing, positioning, styling, connecting, removing, and rendering Excel PivotTable timelines based on date fields
Category: timeline
Primary API: Worksheet.Timelines
Input: A PivotTable whose source includes a valid date field
Output: pivot-timeline.xlsx
Expected Result: A timeline linked to the PivotTable date field is saved at H2.
Product: Aspose.Cells for .NET
Language: C#
*/
```

Keep metadata factual, concise, version-aware, and useful when extracted independently by a RAG system.

## Filename and title rules

Use concise, action-first filenames that express one search intent. Prefer `add-timeline-to-pivottable.cs`. Avoid `example1.cs`, `test.cs`, vague titles, and filenames that encode every implementation step.

## Natural-language opening comment

After metadata, include one sentence stating the operation and expected result:

```csharp
// Add a timeline to a valid PivotTable date field and verify its connection and name.
```

The comment must read like a direct answer, not a keyword list.

## Timelines construction and operation rules

- Create date source data and PivotTable first.
- Resolve the date field index against the PivotTable source.
- Use supported placement and style properties only.
- Refresh/calculate the PivotTable after source/filter changes.
- Verify connection, field, position, and name after reopening.

## Result verification

Verify the resulting timelines object state, relationships, representative values, and artifact. Reopen when persistence is claimed.

An example is incomplete if it performs an operation but never checks the resulting object, value, collection, file, relationship, or rendered artifact.

## Error-handling policy

- Catch only exceptions the scenario can handle meaningfully.
- Include operation and synthetic input context without leaking credentials or workbook data.
- Never suppress failures merely to create an output file.
- Distinguish invalid input, unsupported format/API, corrupt content, unavailable dependencies, and permission failures when possible.
- Let unexpected exceptions fail validation.

## Cache connections, date levels, and filtering

Use documented timeline/cache APIs and verify the affected PivotTables and selected period.

## Layout, style, and rendering

Do not infer chart-like properties such as logarithmic axes unless the Timeline API documents them; rendered output is static.

## Monitoring and interruption

Use documented progress, warning, or interruption APIs only. Keep callbacks lightweight and verify completion or cancellation.

Long-running examples must use version-supported interruption/progress APIs, bounded inputs, cancellation where available, and a verified stopped/completed outcome. Never invent callbacks from task wording.

## Performance and memory examples

Use representative timelines data, batch compatible changes, and report object counts, dimensions, elapsed time, and memory assumptions.

Use `Stopwatch`, identical workloads, warm-up where material, multiple iterations, and report package/framework/environment assumptions. Never present one-machine measurements as universal guarantees.

## Input and output strategy

Prefer generated fixtures. Load existing workbooks only when preserving timelines state is essential. Save to `pivot-timeline.xlsx` and reopen when relevant.

Use relative, deterministic filenames; never developer-specific absolute paths. Do not overwrite inputs unless explicitly requested. Reopen saved output when persistence is part of the claim.

## Security and enterprise safety

Validate untrusted content and identifiers before timelines operations. Bound sizes and avoid logging sensitive values or metadata.

- Never embed licenses, credentials, tokens, personal data, private keys, or connection secrets.
- Keep generated output inside the working directory.
- Treat workbook content and external references as untrusted.

## SEO, GEO, and AEO requirements

### Search intent

Target one primary intent and one or two natural aliases:

- create Excel timeline in C#
- add timeline to PivotTable
- filter PivotTable by date
- remove Excel timeline

Do not stuff every phrase into each example.

### Answer-first structure

The first meaningful comment must identify the operation, primary API, and expected result. An extracted example must reveal what problem is solved, required input, output, and verification without external context.

### Entity consistency

Use canonical names: Aspose.Cells for .NET, C#, Microsoft Excel, Excel workbook, Excel worksheet, Timeline, PivotTable, date field, pivot cache, timeline filter. Avoid ambiguous product nicknames.

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

- [Pivot tables](../pivot-table/)
- [Slicers](../slicer/)
- [PDF](../working-with-pdf/)
- [Charts](../working-with-charts/)

## Definition of done

A `timeline` example is done only when it is technically correct, version-verified, deterministic where possible, safe, runnable, result-checked, clearly named, independently understandable, and retrievable by developers and AI systems.
