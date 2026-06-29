---
name: Aspose.Cells Workbook Management Agent
category: manage-workbook
product: Aspose.Cells for .NET
language: C#
parent: ../AGENTS.md
version: 3.0
last_reviewed: 2026-06-29
primary_intent: Create, configure, inspect, clean, copy, and manage Excel workbooks in C#
primary_apis: [Workbook, WorkbookSettings, WorksheetCollection, ContentTypePropertyCollection, CustomXmlPartCollection]
related_categories: [../open-workbook/, ../save-workbook/, ../working-with-worksheets/, ../document-properties/]
---

# Workbook Management Agent Instructions

## Mission and scope

Create focused enterprise examples for workbook lifecycle and workbook-level structures with Aspose.Cells for .NET. Follow [`../AGENTS.md`](../AGENTS.md).

In scope: workbook creation/lifecycle, settings, worksheet collection coordination, copying between workbooks, content-type properties, custom XML parts, style cleanup, streams, disposal, shared-workbook settings, workbook-wide replace/audit, and batch management.

Use `working-with-worksheets` for worksheet-only operations, `document-properties` for ordinary metadata, `open-workbook` for loading mechanics, and `save-workbook` for output-format behavior.

## Canonical API map

| Intent | APIs |
| --- | --- |
| Create/load workbook | `Workbook` constructors with verified options |
| Workbook settings | `Workbook.Settings` / `WorkbookSettings` |
| Coordinate worksheets | `Workbook.Worksheets` / `WorksheetCollection` |
| Custom XML | `Workbook.CustomXmlParts` and verified collection APIs |
| Content-type metadata | `Workbook.ContentTypeProperties` |
| Remove unused styles | `Workbook.RemoveUnusedStyles` |
| Lifecycle | `IDisposable`, `Dispose`, streams, exception-safe ownership |

## Hard rules

- Keep one workbook-level objective per example.
- Define ownership of workbooks and streams; dispose only objects the example owns.
- Do not dispose a workbook before the caller finishes reading or saving it.
- For streams, document whether they remain open and reset position only when required.
- Verify workbook-wide cleanup does not change visible content or required styles.
- Treat custom XML as untrusted structured input; parse securely and avoid entity expansion.
- Distinguish document properties, content-type properties, and custom XML parts.
- Preserve formulas, styles, drawings, and named ranges deliberately when copying worksheets.

## Canonical pattern

```csharp
using Workbook workbook = new Workbook();
Worksheet worksheet = workbook.Worksheets[0];
worksheet.Name = "Summary";
worksheet.Cells["A1"].PutValue("Managed workbook");

workbook.Save("managed-workbook.xlsx");
```

The root repository generally prefers explicit types. A `using` declaration is acceptable only if repository policy permits it; otherwise use `try/finally` with explicit disposal.

## Example contract

Each example must state workbook source, owned resources, primary workbook operation, affected collections/settings, output, and expected postcondition. Use deterministic workbook content and task-specific output names.

For batch examples, isolate per-file failures, avoid overwriting inputs, bound concurrency, and report successes/failures without leaking paths or workbook data.

## Lifecycle, safety, and performance

- Use memory streams only for genuinely stream-based scenarios.
- Avoid holding many large workbooks simultaneously; dispose each after its output is verified.
- Do not expose custom XML, metadata, or formula content in logs.
- Use secure XML parsing and synthetic data.
- Measure cleanup and batch performance with equivalent inputs and report dimensions/environment.
- Reopen outputs and verify worksheet counts, names, custom XML/content properties, and representative cell/style values.

## Discoverability and validation

Target one intent such as "create Excel workbook in C#," "remove unused Excel styles," "add custom XML to XLSX," or "dispose Aspose.Cells Workbook." The opening comment must identify the workbook-level operation and expected state.

Verify exact collection members and settings against the installed package. Compile, run, reopen, and assert the postcondition. Reject mixed-feature demos, ambiguous ownership, invalid custom XML assumptions, and cleanup claims without comparison.

## Related knowledge

- [Category overview](README.md)
- [Open workbook](../open-workbook/)
- [Save workbook](../save-workbook/)
- [Worksheets](../working-with-worksheets/)
- [Document properties](../document-properties/)
- [Create workbook documentation](https://docs.aspose.com/cells/net/create-new-workbook/)

## Definition of done

The example is done when workbook ownership, scope, operation, affected state, security assumptions, output, and reopened postcondition are explicit and verified.
