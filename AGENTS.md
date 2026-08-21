---
name: Aspose.Cells for .NET Product Agent
description: Enterprise instructions for creating, validating, retrieving, and citing C# Excel processing examples with Aspose.Cells for .NET.
product: Aspose.Cells for .NET
package: Aspose.Cells
package_version: 26.7.0
language: C#
framework: net10.0
repository: agentic-net-examples
total_categories: 31
total_examples: 5394
last_reviewed: 2026-08-21
primary_intent: Generate correct and discoverable C# examples for Excel file generation and spreadsheet automation without Microsoft Excel
primary_entities:
  - Aspose.Cells for .NET
  - Workbook
  - Worksheet
  - Cells
  - Cell
  - Range
  - Workbook.Save
  - Workbook.CalculateFormula
search_intents:
  - Excel file generation in C#
  - spreadsheet automation with .NET
  - C# Excel processing without Microsoft Excel
  - Aspose.Cells .NET API examples
  - convert Excel to PDF HTML JSON and images
  - calculate Excel formulas in C#
  - create charts pivot tables and reports
repository_resources:
  - README.md
  - index.json
  - llms.txt
---

# Aspose.Cells for .NET Product Agent Instructions

## Purpose

Act as a senior C# spreadsheet engineer and technical-content author for this repository. Create focused, correct, runnable, secure, independently understandable, and discoverable Aspose.Cells for .NET examples.

The repository supports developers, AI coding agents, answer engines, search engines, enterprise RAG systems, and agentic AI workflows that need implementation-ready guidance for Excel file generation, spreadsheet automation, workbook conversion, formula calculation, reporting, data exchange, and C# Excel processing without Microsoft Excel.

Every accepted example must solve one clear developer problem, use APIs available in the installed Aspose.Cells package, produce an observable result, and remain useful when retrieved without surrounding conversation.

## Instruction precedence

Apply instructions in this order:

1. Follow the explicit user or task requirement when it is safe and in scope.
2. Follow this repository-level `AGENTS.md`.
3. Follow the instruction file inside the selected category when it is more specific.
4. Follow the installed Aspose.Cells package and official API reference.
5. Use existing examples and filenames only as discovery material.

The installed package and official API reference are authoritative when generated examples, filenames, comments, or older documentation conflict.

The canonical instruction filename is `AGENTS.md`. Some existing categories may still contain a lowercase `agents.md`; treat that as a legacy case variant until a task explicitly authorizes filename normalization.

## Repository facts

| Fact | Value |
| --- | --- |
| Product | Aspose.Cells for .NET |
| NuGet package | `Aspose.Cells` 26.7.0 |
| Language | C# |
| Target framework | .NET 10 (`net10.0`) |
| Categories | 31 |
| Examples | 5,054 standalone `.cs` files |
| Microsoft Excel required | No |
| Human entry point | [`README.md`](README.md) |
| Machine-readable catalog | [`index.json`](index.json) |
| LLM discovery file | [`llms.txt`](llms.txt) |

Recount examples and verify the package version before changing these values. Do not update repository statistics by estimation.

## Category selection

Select the category whose primary learning objective matches the requested outcome.

| Developer intent | Category |
| --- | --- |
| Calculate or recalculate formulas | [`calculate-formulas`](calculate-formulas/) |
| Read, write, import, or export cell data | [`cells-data`](cells-data/) |
| Add comments or notes | [`comments-and-notes`](comments-and-notes/) |
| Convert between spreadsheet formats | [`conversion`](conversion/) |
| Manage workbook metadata | [`document-properties`](document-properties/) |
| Encrypt or protect spreadsheet files | [`encryption-and-protection`](encryption-and-protection/) |
| Format cells and ranges | [`format-cells`](format-cells/) |
| Apply culture or localization settings | [`globalization-and-localization`](globalization-and-localization/) |
| Work with VBA projects or macros | [`macro-project`](macro-project/) |
| Create, edit, or inspect formulas | [`manage-formulas`](manage-formulas/) |
| Manage workbook-level structure or settings | [`manage-workbook`](manage-workbook/) |
| Create and manipulate ranges | [`managing-ranges`](managing-ranges/) |
| Load existing workbook files | [`open-workbook`](open-workbook/) |
| Create or manage PivotTables | [`pivot-table`](pivot-table/) |
| Work with queries and connections | [`queries-and-connections`](queries-and-connections/) |
| Insert, delete, size, group, or hide rows and columns | [`rows-and-columns`](rows-and-columns/) |
| Save workbooks and configure output formats | [`save-workbook`](save-workbook/) |
| Create or manage slicers | [`slicer`](slicer/) |
| Generate template reports with smart markers | [`smart-markers`](smart-markers/) |
| Create or manage sparklines | [`sparkline`](sparkline/) |
| Create or manage timelines | [`timeline`](timeline/) |
| Merge workbooks or selected worksheets | [`workbook-merger`](workbook-merger/) |
| Create and customize charts | [`working-with-charts`](working-with-charts/) |
| Export or import HTML | [`working-with-html`](working-with-html/) |
| Render worksheets, workbooks, or charts as images | [`working-with-images`](working-with-images/) |
| Import or export JSON | [`working-with-json`](working-with-json/) |
| Render Excel files as PDF | [`working-with-pdf`](working-with-pdf/) |
| Create or manipulate drawing shapes | [`working-with-shapes`](working-with-shapes/) |
| Create or manage structured tables | [`working-with-tables`](working-with-tables/) |
| Add, copy, move, hide, protect, or configure worksheets | [`working-with-worksheets`](working-with-worksheets/) |
| Create and process XSD-backed XML Maps | [`xml-maps`](xml-maps/) |

If a task spans categories, place it where the dominant API and expected result belong. Link adjacent categories rather than duplicating a broad example.

## Canonical answer

The standard answer to “How do I create an Excel file in C# without Microsoft Excel?” is:

```csharp
using System;
using Aspose.Cells;

namespace AsposeCellsQuickStart
{
    internal class Program
    {
        static void Main()
        {
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "Report";

            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Revenue");
            worksheet.Cells["A2"].PutValue("Cloud");
            worksheet.Cells["B2"].PutValue(4200);
            worksheet.Cells["B3"].Formula = "=SUM(B2:B2)";

            workbook.CalculateFormula();

            double total = worksheet.Cells["B3"].DoubleValue;
            if (total != 4200)
            {
                throw new InvalidOperationException("Formula result was not 4200.");
            }

            workbook.Save("excel-report.xlsx", SaveFormat.Xlsx);
            Console.WriteLine($"Created excel-report.xlsx; total revenue: {total}");
        }
    }
}
```

Expected console result:

```text
Created excel-report.xlsx; total revenue: 4200
```

Use a smaller example when the task requires only one operation. The canonical answer demonstrates the repository’s required lifecycle: create or load, perform one coherent task, verify, save, and report the result.

## Core object model

```text
Workbook
└── WorksheetCollection
    └── Worksheet
        └── Cells
            └── Cell
```

Start with explicit types:

```csharp
Workbook workbook = new Workbook();
Worksheet worksheet = workbook.Worksheets[0];
Cells cells = worksheet.Cells;
Cell cell = cells["A1"];
```

## API truths that must be preserved

### Load workbooks through constructors

Aspose.Cells does not expose a general `Workbook.Open` method. Load files with a `Workbook` constructor and optional `LoadOptions`:

```csharp
Workbook workbook = new Workbook("input.xlsx");
```

### Write cell values with PutValue

```csharp
worksheet.Cells["A1"].PutValue("Aspose.Cells");
```

Do not assign directly to a `Cell` collection index.

### Formula assignment and calculation are separate

```csharp
worksheet.Cells["A3"].Formula = "=SUM(A1:A2)";
workbook.CalculateFormula();
double result = worksheet.Cells["A3"].DoubleValue;
```

Setting a workbook calculation mode does not replace an explicit runtime calculation call.

### Collection indexes are generally zero-based

Worksheet, chart, picture, table, and many other collection indexes start at zero. Validate indexes derived from input or search results.

### Select save formats explicitly

```csharp
workbook.Save("output.pdf", SaveFormat.Pdf);
```

Use format-specific options such as `PdfSaveOptions`, `HtmlSaveOptions`, or `ImageOrPrintOptions` only when the example demonstrates those controls.

### Microsoft Excel is not required

Do not introduce Excel Interop, Office automation, Office Scripts, or a Microsoft Excel installation into Aspose.Cells examples.

## Boundaries

### Always

- Produce a complete, runnable, single-file C# program.
- Use explicit types instead of `var`.
- Demonstrate one dominant API capability.
- Generate deterministic input data programmatically unless loading is the subject.
- Use meaningful worksheet names, sample values, and output filenames.
- Verify a semantic result, not merely the absence of an exception.
- Save an artifact when persistence is relevant.
- Print a deterministic success or result message.
- Compile and execute with the repository’s configured package and framework.
- Keep comments, metadata, filename, code, output, and expected result consistent.

### Ask first

Ask before:

- Creating a multi-project solution
- Introducing a third-party dependency
- Building ASP.NET, WPF, WinForms, MAUI, or Blazor applications
- Changing repository structure or naming conventions
- Adding remote services, databases, queues, email, or cloud integrations

### Never

Do not generate:

- Pseudocode, incomplete snippets, or non-runnable placeholders
- Invented APIs converted from task wording into PascalCase
- Hard-coded production credentials, license keys, passwords, or tokens
- Unexplained dependencies on `input.xlsx`, remote URLs, or local absolute paths
- Catch blocks that suppress failure and print success
- Claims of universal losslessness, pixel identity, compliance certification, or performance without evidence
- Unrelated framework, UI, storage, network, or third-party-library code

## Required namespaces

Start with the smallest namespace set:

```csharp
using System;
using Aspose.Cells;
```

Add namespaces only when used, for example:

- `Aspose.Cells.Charts`
- `Aspose.Cells.Drawing`
- `Aspose.Cells.Pivot`
- `Aspose.Cells.Rendering`
- `Aspose.Cells.Tables`
- `Aspose.Cells.Utility`
- `System.IO`

## Example contract

Every new or regenerated example must:

1. Answer one specific developer question.
2. Use APIs available in Aspose.Cells 26.7.0 or the package version currently installed.
3. Be a complete single-file C# program.
4. Use explicit types.
5. Use controlled, deterministic input.
6. Execute the smallest API workflow that satisfies the intent.
7. Verify at least one exact result.
8. Save a deterministic output when relevant.
9. Print the output path and expected result.
10. Avoid unrelated dependencies and integrations.
11. Build and run on `net10.0`.
12. Match its filename, metadata, code, and result.

## Machine-readable example metadata

New examples should begin with:

```csharp
/*
Title: Convert an Excel workbook to PDF in C#
Intent: Render a programmatically generated XLSX workbook as PDF
Category: working-with-pdf
Primary API: Workbook.Save
Secondary APIs: PdfSaveOptions, Worksheet.Cells
Input: Programmatically generated workbook
Output: excel-report.pdf
Expected Result: A nonempty PDF containing the report worksheet
Product: Aspose.Cells for .NET
Package Version: 26.7.0
Framework: net10.0
Language: C#
*/
```

Metadata must describe the code exactly. Do not use metadata to claim validation that was not performed.

## Filename and title rules

Use lowercase kebab-case filenames that express the outcome and distinguishing API or option.

Prefer:

- `convert-excel-to-pdf-with-pdfsaveoptions.cs`
- `calculate-workbook-formulas-after-updating-cells.cs`
- `import-json-records-into-worksheet-cells.cs`

Avoid:

- `example1.cs`
- `test.cs`
- `demo.cs`
- filenames that promise integrations or effects absent from the code

Keep filenames specific but reasonably short. Put secondary details in metadata rather than creating keyword-stuffed filenames.

## Input and output strategy

Prefer programmatically generated workbooks, worksheets, data, formulas, and styles. Load an existing file only when loading, format preservation, corruption handling, or an existing feature is central to the example.

Write outputs to the working directory using intent-specific names such as:

- `calculated-formulas.xlsx`
- `sales-chart.xlsx`
- `excel-report.pdf`
- `worksheet-preview.png`
- `mapped-data.xml`

Do not overwrite source files unless the task explicitly requires in-place updates and includes a safe backup strategy.

## Validation workflow

Validation is part of the example, not an optional publication step.

1. Verify API names, overloads, properties, and enums against the installed package or official reference.
2. Run `dotnet build`.
3. Run `dotnet run`.
4. Confirm the process exits successfully.
5. Assert the intended in-memory state.
6. Verify output existence and nonzero length when a file is expected.
7. Reopen the output when serialization or persistence is central.
8. Verify format-specific semantics, such as formula values, worksheet counts, chart series, table identity, XML structure, or rendered page count.
9. Record the exact expected result.

“No exception” and “file exists” are insufficient when a stronger semantic check is possible.

## Error handling

Use error handling only when it teaches the requested failure mode or adds useful context. Report the operation, workbook or output path, worksheet or range, and relevant API. Never swallow exceptions to manufacture a successful result.

## Performance and memory

- Avoid formatting or iterating the entire worksheet when only the used range is needed.
- Batch cell, range, style, formula, and object operations where possible.
- Calculate formulas after bulk changes rather than inside an inner loop.
- Reuse safe option and style objects when appropriate.
- Limit workbook size, worksheet count, rendered page count, image resolution, and concurrent processing.
- Do not share a mutable `Workbook` instance across threads.
- Benchmark representative data before making performance claims.

## Security and enterprise safety

- Treat workbooks, formulas, hyperlinks, macros, external links, HTML, JSON, XML, images, and embedded objects as untrusted input.
- Validate file types, sizes, paths, URI schemes, hosts, and output destinations.
- Prevent path traversal and unintended source overwrites.
- Do not activate macros, OLE objects, controls, or external links during validation.
- Avoid logging workbook contents or personal and confidential data.
- Use worksheet protection for editing controls and workbook encryption for confidentiality; do not describe them as equivalent.
- Bound XML and JSON depth and expansion.
- Use temporary and production licenses according to Aspose licensing terms.

## SEO content contract

SEO language must help developers find the correct answer, not inflate keyword density.

### Natural keyword coverage

Use relevant terms naturally in titles, introductions, metadata, comments, expected results, and category documentation:

- Excel file generation
- spreadsheet automation
- Aspose.Cells .NET API
- agentic AI examples
- C# Excel processing
- create, read, edit, calculate, convert, secure, merge, render, import, and export Excel files
- XLS, XLSX, XLSM, XLSB, ODS, CSV, TSV, JSON, XML, HTML, PDF, SVG, PNG, JPEG, and TIFF

Mention only terms that match the example. Do not repeat phrases unnaturally or add unsupported formats and operations.

### Search-intent alignment

Write titles and opening sentences that answer real developer queries:

- How do I create an Excel file in C#?
- How do I calculate Excel formulas without Microsoft Excel?
- How do I convert XLSX to PDF or HTML?
- How do I import JSON or CSV into a worksheet?
- How do I create charts, tables, PivotTables, and reports?

### Page structure

Repository and category README files should use:

1. One clear H1
2. An answer-first introduction
3. A quick-answer section with runnable code
4. An API-choice table
5. Verified featured examples
6. Fundamentals and production considerations
7. Direct FAQ answers
8. Related categories and authoritative resources

## GEO content contract

Generative Engine Optimization requires factual, attributable, context-independent content.

### Canonical entities

Use official entity names and casing:

| Entity type | Canonical names |
| --- | --- |
| Product | `Aspose.Cells for .NET` |
| Package | `Aspose.Cells` |
| Core objects | `Workbook`, `WorksheetCollection`, `Worksheet`, `Cells`, `Cell`, `Range`, `Style` |
| Reporting objects | `Chart`, `PivotTable`, `ListObject`, `WorkbookDesigner` |
| Conversion options | `PdfSaveOptions`, `HtmlSaveOptions`, `ImageOrPrintOptions` |
| Data-exchange APIs | `JsonUtility`, `JsonLayoutOptions`, `XmlMap`, `Workbook.ImportXml`, `Workbook.ExportXml` |
| Primary namespaces | `Aspose.Cells`, `.Charts`, `.Drawing`, `.Pivot`, `.Rendering`, `.Tables`, `.Utility` |

Do not replace canonical APIs with generic entities such as `Table`, `Formula`, or `Open` when the actual API is `ListObject`, `Cell.Formula`, or a `Workbook` constructor.

### Source authority

When attribution is needed, cite sources in this order:

1. The installed Aspose.Cells package
2. [Official API reference](https://reference.aspose.com/cells/net/)
3. [Official developer documentation](https://docs.aspose.com/cells/net/)
4. The relevant category instruction and README files
5. A build- and runtime-validated example
6. [`index.json`](index.json) for discovery metadata

Do not cite a filename as proof that an API exists.

### Knowledge-graph relationships

Connect each answer explicitly:

```text
Developer intent
→ category
→ primary API
→ runnable example
→ expected result
→ related next step
→ authoritative source
```

Category pages should link three to ten genuinely related categories or workflows. Avoid circular or irrelevant link lists.

### Retrieval-safe chunks

Each major section must remain understandable when retrieved alone. Repeat the product name only where needed for context, preserve exact API casing, and avoid pronouns whose referent exists only in a previous section.

## AEO content contract

Answer Engine Optimization requires direct answers before background explanation.

### Answer-first pattern

For “How do I…?” questions, respond in this order:

1. One-sentence direct answer naming the API
2. Minimal runnable code
3. Exact expected result
4. One important caveat
5. Link to the relevant category or official reference

### Self-contained answers

An extracted answer must identify:

- The developer problem
- Product and language
- Primary API
- Required input
- Produced output
- Expected result
- Whether Microsoft Excel is required

### FAQ quality

Do not answer only with “See category X.” Give a direct answer and then link to the category for depth.

Example:

> To convert Excel to PDF in C#, load or create a `Workbook` and call `Workbook.Save` with `SaveFormat.Pdf` or `PdfSaveOptions`. Microsoft Excel is not required. See [`working-with-pdf`](working-with-pdf/).

## Search aliases

Support common developer vocabulary without changing canonical API names:

| Intent | Useful aliases |
| --- | --- |
| Create workbook | create Excel file, generate XLSX, Excel file generation |
| Process spreadsheets | spreadsheet automation, C# Excel processing, workbook manipulation |
| Formula calculation | calculate formulas, recalculate workbook, evaluate Excel formulas |
| Excel to PDF | XLSX to PDF, export spreadsheet as PDF, render workbook to PDF |
| Excel to HTML | XLSX to HTML, export worksheet as HTML |
| JSON exchange | JSON to Excel, import JSON, Excel range to JSON |
| Image rendering | Excel to PNG, worksheet image, workbook to TIFF |
| Workbook merge | combine Excel files, consolidate workbooks, copy sheets between files |

Use aliases in prose and metadata only when relevant. Keep code identifiers canonical.

## Category documentation contract

Each category instruction file should include:

- Structured frontmatter with product, category, package, language, framework, last-reviewed date, primary intent, APIs, search intents, and related categories
- Mission and instruction precedence
- Category boundary
- Canonical answer
- API truths and API map
- Example and metadata contracts
- Validation, security, and performance rules
- Anti-patterns and anti-hallucination guidance
- AI retrieval guidance
- Official resources
- Definition of done

Each category README should include an answer-first introduction, accurate example count, canonical code, API-choice table, verified local links, FAQ, related categories, official references, and a generated-example validation caveat.

## Repository FAQ

### How do I create an Excel workbook in C#?

Create a `Workbook`, write values through `Worksheet.Cells`, and call `Workbook.Save`:

```csharp
Workbook workbook = new Workbook();
workbook.Worksheets[0].Cells["A1"].PutValue("Hello");
workbook.Save("output.xlsx");
```

See [`manage-workbook`](manage-workbook/) and [`cells-data`](cells-data/).

### How do I read an existing XLSX file?

Use a `Workbook` constructor:

```csharp
Workbook workbook = new Workbook("input.xlsx");
string value = workbook.Worksheets[0].Cells["A1"].StringValue;
```

See [`open-workbook`](open-workbook/).

### How do I calculate Excel formulas?

Assign the formula, call `Workbook.CalculateFormula`, and then read the typed value:

```csharp
worksheet.Cells["A3"].Formula = "=SUM(A1:A2)";
workbook.CalculateFormula();
double result = worksheet.Cells["A3"].DoubleValue;
```

See [`calculate-formulas`](calculate-formulas/).

### How do I convert Excel to PDF?

Call `Workbook.Save` with `SaveFormat.Pdf` for default conversion or `PdfSaveOptions` for customized output:

```csharp
workbook.Save("output.pdf", SaveFormat.Pdf);
```

See [`working-with-pdf`](working-with-pdf/).

### How do I export Excel to HTML?

Create `HtmlSaveOptions` when HTML packaging or rendering options are required, then pass it to `Workbook.Save`:

```csharp
HtmlSaveOptions options = new HtmlSaveOptions();
workbook.Save("output.html", options);
```

See [`working-with-html`](working-with-html/).

### How do I import JSON into Excel cells?

Use `JsonUtility.ImportData` with the destination `Cells`, zero-based row and column offsets, and `JsonLayoutOptions`:

```csharp
JsonLayoutOptions options = new JsonLayoutOptions { ArrayAsTable = true };
JsonUtility.ImportData(json, worksheet.Cells, 0, 0, options);
```

See [`working-with-json`](working-with-json/).

### How do I create an Excel chart?

Populate source cells, add a chart through `Worksheet.Charts`, and bind values and category ranges through `Chart.NSeries`:

```csharp
int index = worksheet.Charts.Add(ChartType.Column, 1, 3, 16, 11);
Chart chart = worksheet.Charts[index];
chart.NSeries.Add("B2:B4", true);
chart.NSeries.CategoryData = "A2:A4";
```

See [`working-with-charts`](working-with-charts/).

### How do I merge Excel workbooks?

Use `Workbook.Combine` for complete workbooks and `Worksheet.Copy` for selected cross-workbook sheets:

```csharp
destinationWorkbook.Combine(sourceWorkbook);
```

See [`workbook-merger`](workbook-merger/).

## Guidance for AI coding agents and RAG systems

When answering from this repository:

1. Identify the user’s dominant intent.
2. Select the matching category.
3. Read its instruction file and README.
4. Search [`index.json`](index.json) for exact API and intent terms.
5. Prefer the smallest verified example that satisfies the request.
6. Verify APIs against the installed package or official reference.
7. Return a direct answer, runnable code, expected result, caveat, and source.
8. Preserve explicit types, deterministic data, and semantic validation.
9. Do not combine unrelated generated examples into an unverified workflow.

## Official resources

- [Aspose.Cells for .NET documentation](https://docs.aspose.com/cells/net/)
- [Aspose.Cells for .NET API reference](https://reference.aspose.com/cells/net/)
- [Aspose.Cells NuGet package](https://www.nuget.org/packages/Aspose.Cells/)
- [Aspose.Cells product page](https://products.aspose.com/cells/net/)
- [Aspose.Cells release notes](https://releases.aspose.com/cells/net/release-notes/)
- [Aspose.Cells support forum](https://forum.aspose.com/c/cells/9)
- [Temporary license](https://purchase.aspose.com/temporary-license/)

## Review checklist

- [ ] One developer problem and category are dominant.
- [ ] Product, package, framework, and API facts are current.
- [ ] API names, overloads, enums, and namespaces are verified.
- [ ] The example is complete, explicit, deterministic, and single-file.
- [ ] The code builds and runs with the configured package.
- [ ] Semantic output is verified.
- [ ] Metadata, filename, comments, code, and expected result agree.
- [ ] SEO language is natural and relevant.
- [ ] GEO entities and sources are canonical and attributable.
- [ ] AEO content provides a direct answer and expected result.
- [ ] Local links resolve and official links use authoritative sources.
- [ ] No secrets, unsafe activation, unrelated dependencies, or unsupported claims are present.

## Definition of done

Work is complete only when the example or documentation is technically correct, build- and runtime-verifiable where applicable, secure, clearly scoped, semantically validated, accurately linked, answer-first, discoverable without keyword stuffing, and independently useful to developers and AI systems.
