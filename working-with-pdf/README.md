---
title: Convert Excel to PDF in C# with Aspose.Cells for .NET
description: Convert XLS, XLSX, XLSM, CSV, and selected Excel worksheets to PDF or PDF/A in C# without Microsoft Excel.
product: Aspose.Cells for .NET
category: working-with-pdf
language: C#
last_reviewed: 2026-08-14
---

# Convert Excel to PDF in C# with Aspose.Cells for .NET

Save Excel workbooks and selected worksheets as PDF with `Workbook.Save`, `SaveFormat.Pdf`, and `PdfSaveOptions`, without Microsoft Excel.

| Repository fact | Value |
| --- | --- |
| Examples | 227 standalone `.cs` files |
| Primary APIs | `Workbook.Save`, `SaveFormat.Pdf`, `PdfSaveOptions`, `PageSetup` |
| Excel required | No |
| Agent guidance | [`agents.md`](agents.md) |
| Catalog | [`../index.json`](../index.json) |

## Quick answer: How do I convert XLSX to PDF in C#?

```csharp
using System;
using System.IO;
using Aspose.Cells;

Workbook workbook = new Workbook();
Worksheet worksheet = workbook.Worksheets[0];
worksheet.Cells["A1"].PutValue("Item");
worksheet.Cells["B1"].PutValue("Quantity");
worksheet.Cells["A2"].PutValue("Apples");
worksheet.Cells["B2"].PutValue(10);

PdfSaveOptions options = new PdfSaveOptions { CalculateFormula = true };
workbook.Save("workbook.pdf", options);

if (new FileInfo("workbook.pdf").Length == 0)
{
    throw new InvalidOperationException("PDF conversion failed.");
}

Console.WriteLine("Saved workbook.pdf.");
```

## Choose the right API

| Need | API |
| --- | --- |
| Default PDF | `Workbook.Save(..., SaveFormat.Pdf)` |
| Customized PDF | `PdfSaveOptions` |
| Margins/orientation/print area | `Worksheet.PageSetup` |
| Selected sheets | `SheetSet` |
| Page subset | `PageIndex`, `PageCount` |
| PDF/A setting | `Compliance` |
| Optimization | `OptimizationType`, `SetImageResample` |

## Featured examples

- [Load XLSX and save as PDF](load-an-xlsx-workbook-from-a-file-path-and-save-it-as-a-pdf-document.cs)
- [Save with configured PdfSaveOptions](save-the-workbook-to-pdf-using-pdfsaveoptions-with-configured-settings-and-verify-output.cs)
- [Calculate formulas before PDF export](use-workbookcalculateformula-to-ensure-all-formulas-are-evaluated-before-exporting-to-pdf.cs)
- [Create one PDF page per worksheet](set-pdfsaveoptionsonepagepersheet-to-true-to-generate-a-separate-pdf-page-for-each-worksheet.cs)
- [Fit all worksheet columns on one page](set-pdfsaveoptionsfitallcolumnsononepage-to-true-to-fit-all-worksheet-columns-onto-a-single-pdf-page.cs)
- [Export selected worksheets](assign-specific-worksheet-indices-to-pdfsaveoptionssheetset-to-export-selected-sheets-as-a-single-pdf.cs)
- [Optimize PDF for minimum size](set-pdfsaveoptionsoptimizationtype-to-minimumsize-for-an-xlsx-workbook-and-save-as-pdf.cs)
- [Configure PDF/A-1a output](set-pdfsaveoptionspdfcompliance-to-pdfa1a-to-produce-pdfa1a-compliant-output-for-archival-purposes.cs)
- [Log font-substitution warnings](retrieve-and-log-font-substitution-warnings-after-rendering-excel-to-pdf.cs)
- [Batch-convert XLSX files to PDF](implement-batch-conversion-of-multiple-xlsx-files-in-a-directory-to-individual-pdf-files.cs)

The corpus is generated. Validate APIs and behavior; avoid suspect scenarios involving PDF parsing, JavaScript, multimedia, unsupported attachments, or third-party validation.

## Getting started

```bash
dotnet new console -n ExcelPdfExample
cd ExcelPdfExample
dotnet add package Aspose.Cells
dotnet build
dotnet run
```

Use `Worksheet.PageSetup` for printed layout. Recalculate formulas before export when values have changed. Ensure required fonts are installed, inspect warnings, and independently validate archival compliance when certification matters.

## FAQ

**Can Aspose.Cells convert Excel to PDF without Excel?** Yes.

**How do I export selected sheets?** Configure a verified `SheetSet` on `PdfSaveOptions`.

**How do I fit one sheet to one page?** Use the appropriate one-page option and understand its scaling effect; fitting all columns is not the same as fitting the entire sheet.

**Does setting PDF/A prove compliance?** No. It configures output; use an independent validator for certification.

**Can Aspose.Cells extract text from an existing PDF?** This category covers generation, not general PDF parsing.

## Related and official resources

- [`conversion`](../conversion/)
- [`working-with-worksheets`](../working-with-worksheets/)
- [`calculate-formulas`](../calculate-formulas/)
- [Excel to PDF](https://docs.aspose.com/cells/net/convert-excel-to-pdf/)
- [PdfSaveOptions](https://reference.aspose.com/cells/net/aspose.cells/pdfsaveoptions/)
- [PageSetup](https://reference.aspose.com/cells/net/aspose.cells/pagesetup/)
- [Aspose.Cells NuGet](https://www.nuget.org/packages/Aspose.Cells/)

Use an Aspose.Cells license or temporary license for full evaluation and review the repository [`LICENSE`](../LICENSE).
