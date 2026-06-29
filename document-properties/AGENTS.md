---
name: Aspose.Cells Document Properties Agent
category: document-properties
product: Aspose.Cells for .NET
language: C#
parent: ../AGENTS.md
version: 3.0
last_reviewed: 2026-06-29
primary_intent: Read, add, update, remove, copy, and audit Excel workbook metadata in C#
primary_apis: [BuiltInDocumentPropertyCollection, CustomDocumentPropertyCollection, DocumentProperty]
related_categories: [../manage-workbook/, ../open-workbook/, ../save-workbook/]
---

# Document Properties Agent Instructions

## Mission and scope

Create focused examples for Excel workbook metadata with Aspose.Cells for .NET. Follow [`../AGENTS.md`](../AGENTS.md) first.

In scope: built-in properties, custom properties, typed metadata, enumeration, existence checks, updates, removal, cloning, persistence, JSON/report export, and metadata governance.

Use `manage-workbook` for custom XML parts and content-type properties unless ordinary document properties remain the main intent.

## Canonical model

| Intent | API |
| --- | --- |
| Standard metadata | `workbook.BuiltInDocumentProperties` |
| User-defined metadata | `workbook.CustomDocumentProperties` |
| One property | `DocumentProperty` |
| Enumerate | Property collection enumerators/indexers verified for the installed version |

## Hard rules

- Choose built-in versus custom properties intentionally.
- Preserve property types such as string, integer, double, Boolean, and `DateTime`.
- Check for an existing custom property before adding one with the same name.
- Treat author, company, manager, comments, and custom values as potentially sensitive metadata.
- Verify persistence by saving and reopening when the task modifies metadata.
- Do not assume every property is supported by every output format.
- Do not confuse document properties with cell values, named ranges, or custom XML parts.

## Canonical pattern

```csharp
Workbook workbook = new Workbook();
workbook.BuiltInDocumentProperties.Title = "Quarterly Report";
workbook.CustomDocumentProperties.Add("ProjectId", 1001);
workbook.Save("workbook-metadata.xlsx");

Workbook reopened = new Workbook("workbook-metadata.xlsx");
string title = reopened.BuiltInDocumentProperties.Title;
Aspose.Cells.Properties.DocumentProperty projectId = reopened.CustomDocumentProperties["ProjectId"];
```

## Example contract

Every example must state the property collection, name, type, operation, and expected persisted value. Use synthetic metadata, explicit types, and deterministic timestamps. Prefer filenames such as `add-custom-document-property-to-xlsx.cs`.

For audits, report property name, type, and a redacted/synthetic value. Never log real author identities, customer names, internal paths, or classification labels.

## Discoverability and validation

Target one intent such as "set Excel title property in C#" or "read custom XLSX metadata." The opening comment must name the property and expected value.

Verify collection members and property conversions against the installed package. Compile, run, reopen output, and compare type and value. Reject examples that stringify all values, overwrite metadata silently, or claim preservation without reopening.

## Related knowledge

- [Category overview](README.md)
- [Manage workbook](../manage-workbook/)
- [Open workbook](../open-workbook/)
- [Official document properties documentation](https://docs.aspose.com/cells/net/managing-document-properties/)

## Definition of done

The example is done when collection, property name/type, privacy treatment, and persisted result are explicit and verified.
