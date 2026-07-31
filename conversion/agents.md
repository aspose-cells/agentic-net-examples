---
name: Aspose.Cells Format Conversion Agent
category: conversion
product: Aspose.Cells for .NET
language: C#
parent: ../AGENTS.md
version: 3.0
last_reviewed: 2026-06-29
primary_intent: Convert Excel workbooks to PDF, HTML, CSV, images, text, ODS, XPS, and other formats in C#
primary_apis: [Workbook.Save, SaveFormat, PdfSaveOptions, HtmlSaveOptions, TxtSaveOptions, ImageOrPrintOptions]
related_categories: [../open-workbook/, ../save-workbook/, ../working-with-pdf/, ../working-with-html/, ../working-with-images/]
---

# File Format Conversion Agent Instructions

## Mission and scope

Create production-minded, reproducible Aspose.Cells for .NET conversion examples. Follow [`../AGENTS.md`](../AGENTS.md), then this guide.

In scope: Excel-to-PDF/HTML/CSV/text/image/ODS/XPS conversions, format-specific save options, worksheet selection, streams, encoding, pagination, fidelity checks, split outputs, and batch conversion.

Use `working-with-pdf`, `working-with-html`, or `working-with-images` when advanced target-specific rendering is the dominant intent. Use `open-workbook` or `save-workbook` when loading or persistence mechanics are primary.

## Canonical API map

| Target | Preferred APIs |
| --- | --- |
| PDF | `Workbook.Save`, `SaveFormat.Pdf`, `PdfSaveOptions` |
| HTML/MHTML | `HtmlSaveOptions`, `SaveFormat.Html`, `SaveFormat.MHtml` |
| CSV/TSV/text | `TxtSaveOptions`, `SaveFormat.Csv`, encoding/separator options |
| Images | `SheetRender`, `WorkbookRender`, `ImageOrPrintOptions` |
| ODS and Excel formats | `Workbook.Save` with an explicit `SaveFormat` or save-options type |
| Streams | Matching `Workbook.Save(Stream, ...)` overload verified for the target |

## Hard rules

- Match file extension, `SaveFormat`, and save-options type.
- State whether conversion applies to the workbook, active sheet, or each sheet.
- Populate formulas and call calculation before conversion when rendered values must be current.
- Configure encoding, delimiter, quoting, and sheet export deliberately for text formats.
- Explain that CSV is single-sheet and does not preserve workbook formatting, charts, formulas, or comments as Excel objects.
- Treat PDF/image conversion as rendering; verify page/image output, not only file existence.
- Use streams safely and reset positions only when the consumer requires it.
- Never claim perfect fidelity without inspecting representative output.

## Canonical pattern

```csharp
Workbook workbook = new Workbook();
Worksheet worksheet = workbook.Worksheets[0];
worksheet.Cells["A1"].PutValue("Revenue");
worksheet.Cells["B1"].PutValue(125000);

PdfSaveOptions options = new PdfSaveOptions
{
    OnePagePerSheet = true
};

workbook.Save("converted-workbook.pdf", options);

if (!File.Exists("converted-workbook.pdf"))
{
    throw new InvalidOperationException("PDF output was not created.");
}
```

Include `System.IO` only when file or stream validation is used.

## Example contract

Each example must identify source and target formats, use one primary conversion objective, configure only relevant options, create deterministic content where possible, and verify output signature/count/size or reopenability.

Metadata must include source, target, primary API, important options, output, and expected result. Prefer filenames such as `convert-xlsx-to-pdf-in-csharp.cs`.

## Fidelity, security, and scale

- Use representative formulas, fonts, charts, images, and page setup only when fidelity is being tested.
- Never embed secrets or personal data in conversion samples.
- Sanitize hyperlinks and HTML; mitigate CSV formula injection for untrusted exported text.
- Bound batch sizes, memory streams, image DPI, and output paths.
- Use fonts available in the deployment environment and report substitution issues.
- For benchmarks, report workbook dimensions, target format, package version, environment, and repeat count.

## Discoverability and validation

Target direct intents such as "convert Excel to PDF in C#" or "export XLSX to CSV with UTF-8." Put source, target, and primary option in the opening comment without keyword stuffing.

Verify overloads and option ownership in the installed package. Compile, run, confirm output exists and is non-empty, then parse/reopen or inspect pages/images where feasible. Reject mismatched extensions, silent sheet loss, and unsupported fidelity claims.

## Related knowledge

- [Category overview](README.md)
- [PDF](../working-with-pdf/)
- [HTML](../working-with-html/)
- [Images](../working-with-images/)
- [Official conversion documentation](https://docs.aspose.com/cells/net/conversion/)

## Definition of done

The example is done when source, target, scope, options, and expected artifact are explicit; output is validated proportionately to the format; and the conversion intent is easy to retrieve and cite.
