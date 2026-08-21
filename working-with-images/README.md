---
title: Convert Excel to Images in C# with Aspose.Cells for .NET
description: Render Excel worksheets, workbooks, and charts to PNG, JPEG, SVG, BMP, and TIFF in C# without Microsoft Excel.
product: Aspose.Cells for .NET
category: working-with-images
language: C#
last_reviewed: 2026-08-14
---

# Convert Excel to Images in C# with Aspose.Cells for .NET

Render Excel worksheets, whole workbooks, and charts to raster or vector images without Microsoft Excel. Use `SheetRender` for worksheet pages, `WorkbookRender` for workbook output, and `Chart.ToImage` for an individual chart.

| Repository fact | Value |
| --- | --- |
| Examples | 57 standalone `.cs` files |
| Primary APIs | `SheetRender`, `WorkbookRender`, `ImageOrPrintOptions`, `Chart.ToImage` |
| Formats | PNG, JPEG, BMP, SVG, TIFF |
| Agent guidance | [`agents.md`](agents.md) |
| Catalog | [`../index.json`](../index.json) |

## Quick answer: How do I convert an Excel worksheet to PNG?

```csharp
using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

Workbook workbook = new Workbook();
Worksheet worksheet = workbook.Worksheets[0];
worksheet.Cells["A1"].PutValue("Revenue");
worksheet.Cells["A2"].PutValue(125000);

ImageOrPrintOptions options = new ImageOrPrintOptions
{
    ImageType = ImageType.Png,
    HorizontalResolution = 150,
    VerticalResolution = 150
};

SheetRender renderer = new SheetRender(worksheet, options);
renderer.ToImage(0, "worksheet.png");

if (new FileInfo("worksheet.png").Length == 0)
{
    throw new InvalidOperationException("PNG rendering failed.");
}

Console.WriteLine($"Rendered {renderer.PageCount} page(s).");
```

## Choose the right API

| Need | API |
| --- | --- |
| Worksheet page to image | `SheetRender.ToImage` |
| Worksheet to TIFF | `SheetRender.ToTiff` |
| Workbook or multipage TIFF | `WorkbookRender` |
| Chart only | `Chart.ToImage` |
| DPI, type, quality, TIFF settings | `ImageOrPrintOptions` |

## Featured examples

- [Convert the first worksheet to PNG](convert-the-first-worksheet-of-a-workbook-to-png-using-default-resolution-for-quick-preview.cs)
- [Save every worksheet as a separate PNG](iterate-through-each-worksheet-in-a-workbook-and-save-each-as-a-separate-png-file.cs)
- [Convert a workbook to multipage TIFF](convert-an-entire-workbook-to-a-multipage-tiff-using-default-rendering-options.cs)
- [Write workbook TIFF output to a memory stream](render-a-workbook-to-tiff-and-write-the-result-into-a-memory-stream-for-further-processing.cs)
- [Convert a worksheet to SVG with a viewBox](convert-a-worksheet-to-svg-with-the-viewbox-attribute-enabled-for-scalable-rendering.cs)
- [Render a worksheet as JPEG with custom quality](generate-a-jpeg-image-from-a-worksheet-with-custom-image-quality-set-to-80-percent.cs)
- [Export a chart to responsive SVG](export-a-chart-to-svg-with-viewbox-attribute-for-responsive-scaling-in-modern-browsers.cs)
- [Preserve chart data labels in PNG output](preserve-data-labels-visibility-when-converting-a-chart-to-png-to-retain-informational-context.cs)

These are generated candidates. Validate each file against the installed Aspose.Cells version before reuse; integration-heavy filenames are not evidence of supported APIs.

## Getting started

```bash
dotnet new console -n ExcelImageExample
cd ExcelImageExample
dotnet add package Aspose.Cells
dotnet build
dotnet run
```

Finalize formulas, styles, charts, fonts, and `PageSetup` before constructing the renderer. Page indexes are zero-based, and one worksheet may render to multiple pages. Higher DPI usually costs more memory and output size. Font availability can change pagination and chart layout.

## FAQ

**Can Aspose.Cells render Excel without Excel?** Yes.

**How do I render only a chart?** Call `Chart.ToImage`.

**How do I create one multipage TIFF?** Use the verified `WorkbookRender` or TIFF workflow for the required scope.

**Why does output differ between servers?** Installed fonts and rendering environment can affect glyphs, sizes, and pagination.

**Is file creation enough validation?** No. Check page count, nonzero length, format, and visual output where fidelity matters.

## Related and official resources

- [`working-with-charts`](../working-with-charts/)
- [`working-with-pdf`](../working-with-pdf/)
- [`working-with-html`](../working-with-html/)
- [Rendering API](https://reference.aspose.com/cells/net/aspose.cells.rendering/)
- [SheetRender](https://reference.aspose.com/cells/net/aspose.cells.rendering/sheetrender/)
- [ImageOrPrintOptions](https://reference.aspose.com/cells/net/aspose.cells.rendering/imageorprintoptions/)
- [Aspose.Cells NuGet](https://www.nuget.org/packages/Aspose.Cells/)

Use an Aspose.Cells license or temporary license for full evaluation and review the repository [`LICENSE`](../LICENSE) before production use.
