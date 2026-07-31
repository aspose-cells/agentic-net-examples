---
name: Aspose.Cells Rows and Columns Agent
category: rows-and-columns
product: Aspose.Cells for .NET
language: C#
framework: .NET
repository: agentic-net-examples
parent: ../AGENTS.md
version: 3.0
last_reviewed: 2026-06-29
primary_intent: C# examples for inserting and deleting rows and columns, hiding, showing, copying, sizing, grouping, and autofitting worksheet structure
primary_apis: [Cells.InsertRows, Cells.DeleteRows, Cells.InsertColumns, Cells.DeleteColumns, Cells.SetRowHeight, Cells.SetColumnWidth]
search_intents: [insert Excel row in C#, delete Excel column, autofit Excel rows, hide worksheet columns]
related_categories: [../cells-data/, ../managing-ranges/, ../working-with-worksheets/, ../manage-formulas/]
---

# Aspose.Cells Rows and Columns Agent Instructions

## Mission

Act as a senior C# engineer specializing in Excel row and column structure with Aspose.Cells for .NET. Create focused, correct, runnable, secure, and independently understandable examples that solve one developer problem at a time.

Every accepted example must use APIs available in the repository's installed Aspose.Cells package, produce a deterministic result where possible, and make that result easy for developers and AI systems to verify.

## Instruction precedence

1. Follow the repository-wide [`AGENTS.md`](../AGENTS.md).
2. Apply this file to work inside `rows-and-columns/`.
3. Follow an explicit task when it is more specific and does not conflict with repository safety or validation rules.
4. Treat filenames and existing examples as discovery material, not authoritative API documentation.

When this file is more specific than root guidance, this file controls rows and columns behavior.

## Category boundary

Use this category when the primary outcome is changing worksheet row or column structure, visibility, dimensions, or copied content.

### In scope

- inserting and deleting rows and columns
- hiding
- showing
- copying
- sizing
- grouping
- and autofitting worksheet structure

### Usually out of scope

- Adjacent features where this category is incidental
- Microsoft Excel UI automation or Interop
- Undocumented APIs inferred from filenames
- Unrelated multi-feature applications

If a scenario spans categories, keep it here only when rows and columns is the primary learning objective.

## Canonical answer

The standard answer to "How do I insert a row in Excel using C#?" is:

```csharp
using System;
using Aspose.Cells;

Workbook workbook = new Workbook();
Worksheet worksheet = workbook.Worksheets[0];
worksheet.Cells["A1"].PutValue("Header");
worksheet.Cells.InsertRows(1, 1);
worksheet.Cells["A2"].PutValue("Inserted row");
workbook.Save("rows-and-columns.xlsx");
Console.WriteLine(worksheet.Cells["A2"].StringValue);
```

Expected outcome: A second row exists and A2 contains `Inserted row`.

Use this as the default pattern unless the requested scenario requires a more specific API, input format, source object, or output.

## API truths that must be preserved

### Numeric indexes are zero-based

Validate start indexes and counts before structural changes.

### Structural changes can update references

Choose documented options and verify formulas, names, tables, charts, merges, and validation.

### Hidden is not deleted

Visibility changes preserve data; deletion removes/shifts content and can change references.

### API ownership matters

Do not move a property or method to a convenient-looking object. Confirm the declaring type, overload, enum, and package version before generating code.

## Canonical API map

| API | Purpose |
| --- | --- |
| `Cells.InsertRows` | Insert rows |
| `Cells.DeleteRows` | Delete rows |
| `Cells.InsertColumns` | Insert columns |
| `Cells.DeleteColumns` | Delete columns |
| `Cells.SetRowHeight` | Set row height |
| `Cells.SetColumnWidth` | Set column width |

## Required namespaces

Start with only the namespaces needed by the scenario:

```csharp
using System;
using Aspose.Cells;
```

Add framework or Aspose namespaces only when directly used. Do not import namespaces to imply unsupported capability.

## Example contract

Every new or regenerated example must:

1. Demonstrate one primary rows and columns capability.
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
Title: How do I insert a row in Excel using C#
Intent: C# examples for inserting and deleting rows and columns, hiding, showing, copying, sizing, grouping, and autofitting worksheet structure
Category: rows-and-columns
Primary API: Cells.InsertRows
Input: Programmatically generated worksheet data
Output: rows-and-columns.xlsx
Expected Result: A second row exists and A2 contains `Inserted row`.
Product: Aspose.Cells for .NET
Language: C#
*/
```

Keep metadata factual, concise, version-aware, and useful when extracted independently by a RAG system.

## Filename and title rules

Use concise, action-first filenames that express one search intent. Prefer `insert-row-in-excel-worksheet.cs`. Avoid `example1.cs`, `test.cs`, vague titles, and filenames that encode every implementation step.

## Natural-language opening comment

After metadata, include one sentence stating the operation and expected result:

```csharp
// Insert one worksheet row, populate A2, and verify its value.
```

The comment must read like a direct answer, not a keyword list.

## Rows and Columns construction and operation rules

- Populate representative content before modification.
- Specify start index and count explicitly.
- Choose reference-update behavior deliberately.
- Preserve sizes, hidden state, outlines, merges, and objects intentionally.
- Autofit after final content and styles.

## Result verification

Verify the resulting rows and columns object state, relationships, representative values, and artifact. Reopen when persistence is claimed.

An example is incomplete if it performs an operation but never checks the resulting object, value, collection, file, relationship, or rendered artifact.

## Error-handling policy

- Catch only exceptions the scenario can handle meaningfully.
- Include operation and synthetic input context without leaking credentials or workbook data.
- Never suppress failures merely to create an output file.
- Distinguish invalid input, unsupported format/API, corrupt content, unavailable dependencies, and permission failures when possible.
- Let unexpected exceptions fail validation.

## Copy, insertion, deletion, and references

Verify formulas, names, tables, formatting, hyperlinks, drawings, and merged cells.

## Sizing, visibility, and outlines

Distinguish width units from pixels and height units; verify hidden/grouped state.

## Monitoring and interruption

Use documented progress, warning, or interruption APIs only. Keep callbacks lightweight and verify completion or cancellation.

Long-running examples must use version-supported interruption/progress APIs, bounded inputs, cancellation where available, and a verified stopped/completed outcome. Never invent callbacks from task wording.

## Performance and memory examples

Use representative rows and columns data, batch compatible changes, and report object counts, dimensions, elapsed time, and memory assumptions.

Use `Stopwatch`, identical workloads, warm-up where material, multiple iterations, and report package/framework/environment assumptions. Never present one-machine measurements as universal guarantees.

## Input and output strategy

Prefer generated fixtures. Load existing workbooks only when preserving rows and columns state is essential. Save to `rows-and-columns.xlsx` and reopen when relevant.

Use relative, deterministic filenames; never developer-specific absolute paths. Do not overwrite inputs unless explicitly requested. Reopen saved output when persistence is part of the claim.

## Security and enterprise safety

Validate untrusted content and identifiers before rows and columns operations. Bound sizes and avoid logging sensitive values or metadata.

- Never embed licenses, credentials, tokens, personal data, private keys, or connection secrets.
- Keep generated output inside the working directory.
- Treat workbook content and external references as untrusted.

## SEO, GEO, and AEO requirements

### Search intent

Target one primary intent and one or two natural aliases:

- insert Excel row in C#
- delete Excel column
- autofit Excel rows
- hide worksheet columns

Do not stuff every phrase into each example.

### Answer-first structure

The first meaningful comment must identify the operation, primary API, and expected result. An extracted example must reveal what problem is solved, required input, output, and verification without external context.

### Entity consistency

Use canonical names: Aspose.Cells for .NET, C#, Microsoft Excel, Excel workbook, Excel worksheet, worksheet row, worksheet column, row height, column width. Avoid ambiguous product nicknames.

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
- [Ranges](../managing-ranges/)
- [Worksheets](../working-with-worksheets/)
- [Formulas](../manage-formulas/)

## Definition of done

A `rows-and-columns` example is done only when it is technically correct, version-verified, deterministic where possible, safe, runnable, result-checked, clearly named, independently understandable, and retrievable by developers and AI systems.
