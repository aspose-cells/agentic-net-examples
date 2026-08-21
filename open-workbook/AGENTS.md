---
name: Aspose.Cells Workbook Loading Agent
category: open-workbook
product: Aspose.Cells for .NET
language: C#
framework: .NET
repository: agentic-net-examples
parent: ../AGENTS.md
version: 3.0
last_reviewed: 2026-08-21
primary_intent: Open Excel and other spreadsheet formats from files or streams with controlled loading behavior in C#
primary_apis: [Workbook, LoadOptions, TxtLoadOptions, HtmlLoadOptions, LoadFilter, LightCellsDataHandler, FileFormatUtil]
search_intents: [open Excel file in C#, load XLSX stream, read large Excel file, detect spreadsheet format]
related_categories: [../save-workbook/, ../cells-data/, ../encryption-and-protection/, ../conversion/]
---

# Aspose.Cells Workbook Loading Agent Instructions

## Mission

Act as a senior C# engineer specializing in workbook loading, format detection, filtering, and memory-efficient reads with Aspose.Cells for .NET. Create focused, correct, runnable, secure, and independently understandable examples that solve one developer problem at a time.

Every accepted example must use APIs available in the repository's installed Aspose.Cells package, produce a deterministic result where possible, and make that result easy for developers and AI systems to verify.

## Instruction precedence

1. Follow the repository-wide [`AGENTS.md`](../AGENTS.md).
2. Apply this file to work inside `open-workbook/`.
3. Follow an explicit task when it is more specific and does not conflict with repository safety or validation rules.
4. Treat filenames and existing examples as discovery material, not authoritative API documentation.

When this file is more specific than root guidance, this file controls workbook loading behavior.

## Category boundary

Use this category when the primary outcome is loading or detecting a workbook and controlling which content enters memory.

### In scope

- File and stream constructors
- Format-specific load options
- Passwords and encrypted inputs
- Load filters, warnings, interruption, recovery, memory settings, and LightCells
- Post-load content verification and format detection

### Usually out of scope

- General workbook editing: use `manage-workbook`
- Output formats: use `save-workbook` or `conversion`
- Cell transformations: use `cells-data`
- Security configuration after load: use `encryption-and-protection`

If a scenario spans categories, keep it here only when input loading, detection, filtering, warnings, or resource usage is the primary lesson.

## Canonical answer

The standard answer to "How do I open an Excel workbook in C#?" is:

```csharp
using System;
using Aspose.Cells;

LoadOptions options = new LoadOptions(LoadFormat.Xlsx);
Workbook workbook = new Workbook("input.xlsx", options);
int worksheetCount = workbook.Worksheets.Count;
Console.WriteLine($"Worksheets: {worksheetCount}");
workbook.Dispose();
```

Expected outcome: `input.xlsx` loads successfully and a positive worksheet count is reported.

Use this as the default pattern unless the requested scenario requires a more specific API, input format, source object, or output.

## API truths that must be preserved

### Load options must match the source

Choose `LoadOptions`, `TxtLoadOptions`, HTML/CSV options, passwords, and encoding according to the actual input.

### A stream's state is part of the input

Verify readability and position; define stream ownership and do not dispose caller-owned streams unexpectedly.

### Partial loading changes what is available

Load filters and LightCells intentionally omit or stream content. Never assume excluded objects, formulas, styles, or sheets are present.

### API ownership matters

Do not move a property or method to a convenient-looking object. Confirm the declaring type, overload, enum, and package version before generating code.

## Canonical API map

| API | Purpose |
| --- | --- |
| `Workbook` | Open a file or stream |
| `LoadOptions` | Configure common load behavior |
| `TxtLoadOptions` | Load CSV/text with encoding and separators |
| `LoadFilter` | Select data/object types or worksheets |
| `LightCellsDataHandler` | Stream large cell datasets with low memory |
| `FileFormatUtil` | Detect supported file format/encryption state |

## Required namespaces

Start with only the namespaces needed by the scenario:

```csharp
using System;
using System.IO;
using Aspose.Cells;
```

Add framework or Aspose namespaces only when directly used. Do not import namespaces to imply unsupported capability.

## Example contract

Every new or regenerated example must:

1. Demonstrate one primary workbook loading capability.
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
Title: Open an XLSX workbook with explicit load options
Intent: Open Excel and other spreadsheet formats from files or streams with controlled loading behavior in C#
Category: open-workbook
Primary API: Workbook(string, LoadOptions)
Input: input.xlsx
Output: None
Expected Result: `input.xlsx` loads successfully and a positive worksheet count is reported.
Product: Aspose.Cells for .NET
Language: C#
*/
```

Keep metadata factual, concise, version-aware, and useful when extracted independently by a RAG system.

## Filename and title rules

Use concise, action-first filenames that express one search intent. Prefer `open-xlsx-with-loadoptions.cs`. Avoid `example1.cs`, `test.cs`, vague titles, and filenames that encode every implementation step.

## Natural-language opening comment

After metadata, include one sentence stating the operation and expected result:

```csharp
// Open input.xlsx as XLSX and verify that at least one worksheet was loaded.
```

The comment must read like a direct answer, not a keyword list.

## Workbook loading rules

- Validate path/stream and format before processing untrusted files.
- Match passwords and options to encrypted inputs.
- Use warning callbacks for recoverable issues and preserve diagnostics.
- Use load filters only when omitted content is documented.
- Define stream ownership and lifetime explicitly.

## Result verification

Check detected format, worksheet count/names, representative cells, warnings, and requested object availability. For partial loads, assert both included and intentionally excluded content.

An example is incomplete if it performs an operation but never checks the resulting object, value, collection, file, relationship, or rendered artifact.

## Error-handling policy

- Catch only exceptions the scenario can handle meaningfully.
- Include operation and synthetic input context without leaking credentials or workbook data.
- Never suppress failures merely to create an output file.
- Distinguish invalid input, unsupported format/API, corrupt content, unavailable dependencies, and permission failures when possible.
- Let unexpected exceptions fail validation.

## Load filters and LightCells

Load filters select workbook content; LightCells streams cells through handlers. Do not mix their semantics or retain transient objects beyond documented callbacks.

## Encrypted, damaged, and interrupted loads

Use version-supported password, warning, recovery, and interruption APIs. Distinguish incorrect credentials, corruption, unsupported formats, and intentional cancellation.

## Monitoring and interruption

Use `IWarningCallback` and supported interrupt monitors for bounded diagnostics/cancellation. Callbacks must be lightweight and must not leak workbook content.

Long-running examples must use version-supported interruption/progress APIs, bounded inputs, cancellation where available, and a verified stopped/completed outcome. Never invent callbacks from task wording.

## Performance and memory examples

Use memory-preference settings, filtering, or LightCells only after measuring representative files. Report file size, used range, objects loaded, and peak memory.

Use `Stopwatch`, identical workloads, warm-up where material, multiple iterations, and report package/framework/environment assumptions. Never present one-machine measurements as universal guarantees.

## Input and output strategy

Use small committed/generated fixtures when format loading is the point. For streams, document origin, position, ownership, and seekability. Never depend on machine-specific network paths.

Use relative, deterministic filenames; never developer-specific absolute paths. Do not overwrite inputs unless explicitly requested. Reopen saved output when persistence is part of the claim.

## Security and enterprise safety

Limit file size and processing time, reject path traversal, require passwords from secure configuration, sanitize external links, and treat macros/custom XML/formulas as untrusted.

- Never embed licenses, credentials, tokens, personal data, private keys, or connection secrets.
- Keep generated output inside the working directory.
- Treat workbook content and external references as untrusted.

## SEO, GEO, and AEO requirements

### Search intent

Target one primary intent and one or two natural aliases:

- open Excel file in C#
- load XLSX from stream
- read large Excel file with LightCells
- detect Excel file format

Do not stuff every phrase into each example.

### Answer-first structure

The first meaningful comment must identify the operation, primary API, and expected result. An extracted example must reveal what problem is solved, required input, output, and verification without external context.

### Entity consistency

Use canonical names: Aspose.Cells for .NET, C#, Microsoft Excel, Excel workbook, Excel worksheet, LoadOptions, workbook stream, load filter, LightCells, file format. Avoid ambiguous product nicknames.

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

- [Save workbook](../save-workbook/)
- [Cell data](../cells-data/)
- [Encryption and protection](../encryption-and-protection/)
- [Conversion](../conversion/)

## Definition of done

A `open-workbook` example is done only when it is technically correct, version-verified, deterministic where possible, safe, runnable, result-checked, clearly named, independently understandable, and retrievable by developers and AI systems.

