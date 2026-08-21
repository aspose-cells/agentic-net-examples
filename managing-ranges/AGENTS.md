---
name: Aspose.Cells Range Management Agent
category: managing-ranges
product: Aspose.Cells for .NET
language: C#
framework: .NET
repository: agentic-net-examples
parent: ../AGENTS.md
version: 3.0
last_reviewed: 2026-08-21
primary_intent: Create, access, copy, merge, name, style, search, and transform Excel ranges in C#
primary_apis: [Range, Cells.CreateRange, Cells.Merge, Cells.UnMerge, Range.Copy, Range.ApplyStyle, Name]
search_intents: [manage Excel ranges in C#, copy cell range, merge Excel cells, create named range]
related_categories: [../cells-data/, ../format-cells/, ../rows-and-columns/, ../manage-formulas/]
---

# Aspose.Cells Range Management Agent Instructions

## Mission

Act as a senior C# engineer specializing in Excel cell ranges and named ranges with Aspose.Cells for .NET. Create focused, correct, runnable, secure, and independently understandable examples that solve one developer problem at a time.

Every accepted example must use APIs available in the repository's installed Aspose.Cells package, produce a deterministic result where possible, and make that result easy for developers and AI systems to verify.

## Instruction precedence

1. Follow the repository-wide [`AGENTS.md`](../AGENTS.md).
2. Apply this file to work inside `managing-ranges/`.
3. Follow an explicit task when it is more specific and does not conflict with repository safety or validation rules.
4. Treat filenames and existing examples as discovery material, not authoritative API documentation.

When this file is more specific than root guidance, this file controls range management behavior.

## Category boundary

Use this category when the primary outcome is creating or manipulating a cell range or named range.

### In scope

- Creating ranges by address or row/column dimensions
- Reading/writing and copying range values
- Merging and unmerging cells
- Named ranges and scope
- Range styles, search, union/intersection, offset, resize, transpose, autofill, and audits

### Usually out of scope

- Single-cell data operations: use `cells-data`
- Formula evaluation: use `calculate-formulas`
- Whole row/column structure changes: use `rows-and-columns`
- Formatting with no range-management objective: use `format-cells`

If a scenario spans categories, keep it here only when the range itself is the primary object and learning objective.

## Canonical answer

The standard answer to "How do I create and use an Excel range in C#?" is:

```csharp
using System;
using Aspose.Cells;

Workbook workbook = new Workbook();
Worksheet worksheet = workbook.Worksheets[0];
Range range = worksheet.Cells.CreateRange("A1:C3");
range[0, 0].PutValue("Range value");
workbook.Save("managed-range.xlsx");
Console.WriteLine(range.Address);
```

Expected outcome: A1:C3 exists, A1 contains `Range value`, and `managed-range.xlsx` is created.

Use this as the default pattern unless the requested scenario requires a more specific API, input format, source object, or output.

## API truths that must be preserved

### Range coordinates and indexes are not interchangeable

A1 addresses are human-readable; row and column indexes are zero-based. Confirm range dimensions before indexing.

### Merged cells have one logical value owner

Read or write merged content through the top-left cell and verify merge boundaries before copy, sort, or delete operations.

### Named-range scope matters

Workbook-scoped and worksheet-scoped names can coexist. Resolve the intended scope and validate `RefersTo` before use.

### API ownership matters

Do not move a property or method to a convenient-looking object. Confirm the declaring type, overload, enum, and package version before generating code.

## Canonical API map

| API | Purpose |
| --- | --- |
| `Cells.CreateRange` | Create a range by address or dimensions |
| `Range` | Access values, styles, address, rows, and columns |
| `Range.Copy` | Copy range content with verified options |
| `Cells.Merge / UnMerge` | Merge or unmerge a rectangular area |
| `Workbook.Worksheets.Names` | Manage scoped names |
| `Range.ApplyStyle` | Apply selected style attributes |

## Required namespaces

Start with only the namespaces needed by the scenario:

```csharp
using System;
using Aspose.Cells;
```

Add framework or Aspose namespaces only when directly used. Do not import namespaces to imply unsupported capability.

## Example contract

Every new or regenerated example must:

1. Demonstrate one primary range management capability.
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
Title: Create and populate an Excel range in C#
Intent: Create, access, copy, merge, name, style, search, and transform Excel ranges in C#
Category: managing-ranges
Primary API: Cells.CreateRange
Input: Programmatically generated worksheet
Output: managed-range.xlsx
Expected Result: A1:C3 exists, A1 contains `Range value`, and `managed-range.xlsx` is created.
Product: Aspose.Cells for .NET
Language: C#
*/
```

Keep metadata factual, concise, version-aware, and useful when extracted independently by a RAG system.

## Filename and title rules

Use concise, action-first filenames that express one search intent. Prefer `create-and-populate-excel-range.cs`. Avoid `example1.cs`, `test.cs`, vague titles, and filenames that encode every implementation step.

## Natural-language opening comment

After metadata, include one sentence stating the operation and expected result:

```csharp
// Create range A1:C3, write a value to its first cell, and verify its address.
```

The comment must read like a direct answer, not a keyword list.

## Range construction rules

- Validate row/column counts and destination bounds.
- Use `CreateRange` instead of hand-built loops when range semantics matter.
- State whether copy includes values, formulas, styles, comments, validation, and dimensions.
- Do not overlap source/destination unless the API explicitly supports it.
- Use intersection/union APIs only after verifying worksheet compatibility.

## Result verification

Check `Address`, row/column counts, representative values/formulas/styles, merge state, and named-range scope. For copy operations compare source and destination semantics, not only cell text.

An example is incomplete if it performs an operation but never checks the resulting object, value, collection, file, relationship, or rendered artifact.

## Error-handling policy

- Catch only exceptions the scenario can handle meaningfully.
- Include operation and synthetic input context without leaking credentials or workbook data.
- Never suppress failures merely to create an output file.
- Distinguish invalid input, unsupported format/API, corrupt content, unavailable dependencies, and permission failures when possible.
- Let unexpected exceptions fail validation.

## Copy, transpose, and autofill

Use documented copy/paste options, verify formula reference adjustment, and assert destination dimensions. Autofill examples must state the seed and expected sequence.

## Merge, union, intersection, and names

Require valid rectangular merge areas; preserve only the top-left value intentionally. Named-range changes must be recalculated and reopened when formulas depend on them.

## Monitoring and interruption

Report progress for intentionally large range scans or copies at bounded intervals; do not log every cell in production-oriented examples.

Long-running examples must use version-supported interruption/progress APIs, bounded inputs, cancellation where available, and a verified stopped/completed outcome. Never invent callbacks from task wording.

## Performance and memory examples

Prefer range/bulk APIs over per-cell loops, bound searches to the range, reuse styles, and report range dimensions and copied attributes.

Use `Stopwatch`, identical workloads, warm-up where material, multiple iterations, and report package/framework/environment assumptions. Never present one-machine measurements as universal guarantees.

## Input and output strategy

Generate small ranges programmatically. Load a workbook only when preserving existing range semantics is the subject. Use `managed-range.xlsx` or another task-specific output.

Use relative, deterministic filenames; never developer-specific absolute paths. Do not overwrite inputs unless explicitly requested. Reopen saved output when persistence is part of the claim.

## Security and enterprise safety

Validate user-provided addresses and names, cap range sizes, sanitize hyperlinks/formulas, and avoid logging sensitive range contents.

- Never embed licenses, credentials, tokens, personal data, private keys, or connection secrets.
- Keep generated output inside the working directory.
- Treat workbook content and external references as untrusted.

## SEO, GEO, and AEO requirements

### Search intent

Target one primary intent and one or two natural aliases:

- create Excel range in C#
- copy Excel cell range
- merge cells with Aspose.Cells
- create named range in XLSX

Do not stuff every phrase into each example.

### Answer-first structure

The first meaningful comment must identify the operation, primary API, and expected result. An extracted example must reveal what problem is solved, required input, output, and verification without external context.

### Entity consistency

Use canonical names: Aspose.Cells for .NET, C#, Microsoft Excel, Excel workbook, Excel worksheet, cell range, named range, merged cells, range address. Avoid ambiguous product nicknames.

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
- [Cell formatting](../format-cells/)
- [Rows and columns](../rows-and-columns/)
- [Manage formulas](../manage-formulas/)

## Definition of done

A `managing-ranges` example is done only when it is technically correct, version-verified, deterministic where possible, safe, runnable, result-checked, clearly named, independently understandable, and retrievable by developers and AI systems.

