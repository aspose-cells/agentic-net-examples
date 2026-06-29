---
title: Create Excel Charts in C# with Aspose.Cells for .NET
description: Create, bind, customize, inspect, and export Excel charts in C# without Microsoft Excel.
product: Aspose.Cells for .NET
category: working-with-charts
language: C#
last_reviewed: 2026-06-29
---

# Create Excel Charts in C# with Aspose.Cells for .NET

Use `Worksheet.Charts`, `Chart`, and `Chart.NSeries` to create and customize Excel charts without Microsoft Excel.

| Repository fact | Value |
| --- | --- |
| Examples | 521 standalone `.cs` files |
| Primary APIs | `ChartCollection`, `Chart`, `SeriesCollection`, `Series`, `Axis` |
| Excel required | No |
| Agent guidance | [`agents.md`](agents.md) |
| Catalog | [`../index.json`](../index.json) |

## Quick answer: How do I create an Excel chart in C#?

```csharp
using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

Workbook workbook = new Workbook();
Worksheet worksheet = workbook.Worksheets[0];
worksheet.Cells["A1"].PutValue("Quarter");
worksheet.Cells["B1"].PutValue("Revenue");
worksheet.Cells["A2"].PutValue("Q1");
worksheet.Cells["A3"].PutValue("Q2");
worksheet.Cells["B2"].PutValue(120);
worksheet.Cells["B3"].PutValue(180);

int index = worksheet.Charts.Add(ChartType.Column, 1, 3, 16, 11);
Chart chart = worksheet.Charts[index];
chart.Title.Text = "Quarterly Revenue";
chart.NSeries.Add("B2:B3", true);
chart.NSeries.CategoryData = "A2:A3";

if (chart.NSeries.Count != 1)
{
    throw new InvalidOperationException("Series binding failed.");
}

workbook.Save("column-chart.xlsx");
Console.WriteLine("Created column-chart.xlsx with one chart.");
```

## API choice

| Need | API |
| --- | --- |
| Create chart | `Worksheet.Charts.Add` |
| Bind values/categories | `Chart.NSeries` |
| Customize series | `Series` |
| Configure axes | `Axis` |
| Add labels | `DataLabels` |
| Export chart image/PDF | `Chart.ToImage`, `Chart.ToPdf` |

## Featured examples

- [Create a column chart](create-a-column-chart-object-on-the-worksheet-to-visualize-the-sales-data.cs)
- [Add a pie chart](add-a-pie-chart-object-to-the-worksheet-using-charttypepie.cs)
- [Create a line chart](create-a-line-chart-on-a-worksheet-using-data-from-cells-a1-through-a10.cs)
- [Create a combo chart with a secondary axis](construct-a-combo-chart-that-combines-a-column-series-with-a-line-series-sharing-a-secondary-axis.cs)
- [Export a chart to 300-DPI PNG](export-a-specific-chart-to-a-png-image-file-with-a-resolution-of-three-hundred-dpi.cs)
- [Validate labels after changing source data](validate-that-data-labels-display-correct-cell-values-after-modifying-the-source-cell-range-programmatically.cs)
- [Verify pie-slice colors](verify-custom-slice-colors-by-comparing-chartpointforegroundcolor-to-expected-rgb-values.cs)
- [Write chart PDF output to memory](write-the-chart-pdf-output-to-a-memorystream-for-further-inmemory-processing.cs)

Generated chart examples can be specialized or version-sensitive. Compile, run, and inspect their source ranges and output.

## Getting started

```bash
dotnet new console -n ExcelChartExample
cd ExcelChartExample
dotnet add package Aspose.Cells
dotnet build
dotnet run
```

Populate source data first. Add the chart, retrieve it by returned index, bind values and categories separately, configure it, verify semantic state, and save. Recalculate formula-fed data before rendering.

## FAQ

**What is `NSeries`?** It is the Aspose.Cells series collection on `Chart`.

**How do I add a secondary axis?** Put a compatible series on the secondary axis with the verified series property.

**How do I export only the chart?** Use `Chart.ToImage` or `Chart.ToPdf`.

**Why does chart rendering vary?** Fonts and rendering environment can affect layout.

## Related and official resources

- [`cells-data`](../cells-data/)
- [`calculate-formulas`](../calculate-formulas/)
- [`working-with-images`](../working-with-images/)
- [Charts documentation](https://docs.aspose.com/cells/net/charts/)
- [Chart API](https://reference.aspose.com/cells/net/aspose.cells.charts/chart/)
- [ChartCollection](https://reference.aspose.com/cells/net/aspose.cells.charts/chartcollection/)
- [Aspose.Cells NuGet](https://www.nuget.org/packages/Aspose.Cells/)

Review the repository [`LICENSE`](../LICENSE) and Aspose licensing terms before production use.
