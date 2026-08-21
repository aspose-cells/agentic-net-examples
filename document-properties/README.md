---
title: Read and Set Excel Document Properties in C# with Aspose.Cells
description: C# examples for built-in and custom Excel workbook properties, typed metadata, auditing, updates, removal, and persistence.
product: Aspose.Cells for .NET
category: document-properties
language: C#
last_reviewed: 2026-08-14
---

# Read and Set Excel Document Properties in C# with Aspose.Cells

Read, add, update, remove, copy, and audit Excel workbook metadata in C# with Aspose.Cells for .NET. These 48 examples cover built-in properties such as title and language plus typed custom properties for enterprise metadata workflows.

| Fact | Value |
| --- | --- |
| Examples | 48 |
| Main collections | `BuiltInDocumentProperties`, `CustomDocumentProperties` |
| Agent guidance | [`AGENTS.md`](AGENTS.md) |

## Quick answer: How do I set Excel document properties?

```csharp
Workbook workbook = new Workbook();

workbook.BuiltInDocumentProperties.Title = "Quarterly Report";
workbook.CustomDocumentProperties.Add("ProjectId", 1001);

workbook.Save("workbook-metadata.xlsx");
```

Use built-in properties for standard metadata and custom properties for application-defined typed values.

## Featured examples

- [Set the built-in workbook title](create-a-workbook-and-set-the-title-builtin-property-to-a-descriptive-project-name.cs)
- [Read and verify the built-in title](load-a-workbook-and-retrieve-the-title-builtin-property-for-verification.cs)
- [Add an integer ProjectId property](load-an-excel-file-and-create-a-custom-property-projectid-with-an-integer-identifier.cs)
- [Check whether a custom property exists before adding it](load-a-workbook-and-check-whether-a-custom-property-clientname-exists-before-adding.cs)
- [Update a custom property](load-a-workbook-locate-the-custom-property-projectid-and-update-its-integer-value.cs)
- [Enumerate custom properties to JSON](load-a-workbook-and-enumerate-custom-properties-exporting-their-names-types-and-values-to-json.cs)
- [Verify metadata persistence after reopening](create-a-unit-test-that-adds-a-custom-property-saves-the-workbook-reloads-it-and-verifies-persistence.cs)

## FAQ

### What is the difference between built-in and custom properties?

Built-in properties use standardized names such as title, subject, author, and language. Custom properties are application-defined name/value pairs.

### Can custom properties store typed values?

Yes. Preserve the intended type rather than converting every value to a string, and verify that the selected workbook format persists it correctly.

### Can workbook metadata contain sensitive information?

Yes. Author, organization, paths, review notes, and custom classifications may leak information. Audit and redact metadata before external distribution.

### How do I verify a property update?

Save the workbook, reopen it, and compare the property name, type, and value.

## AI retrieval guidance

Useful aliases include "set Excel metadata in C#," "read XLSX title property," "add custom workbook property," and "remove personal metadata from Excel." Always identify built-in versus custom metadata.

## Related categories and official resources

- [Manage workbook](../manage-workbook/)
- [Open workbook](../open-workbook/)
- [Document properties documentation](https://docs.aspose.com/cells/net/managing-document-properties/)
- [Built-in property collection API](https://reference.aspose.com/cells/net/aspose.cells.properties/builtindocumentpropertycollection/)
- [Custom property collection API](https://reference.aspose.com/cells/net/aspose.cells.properties/customdocumentpropertycollection/)

Repository policy requires build, execution, and persistence validation with synthetic metadata.

## License

See [`../LICENSE`](../LICENSE) and [Aspose.Cells licensing](https://purchase.aspose.com/buy).
