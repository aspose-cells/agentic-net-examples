---
name: Aspose.Cells Charts Agent
category: working-with-charts
product: Aspose.Cells for .NET
language: C#
framework: .NET
repository: agentic-net-examples
parent: ../AGENTS.md
version: 3.0
last_reviewed: 2026-08-21
primary_intent: Create, customize, inspect, and render Excel charts in C#
primary_apis: [Worksheet.Charts, ChartCollection.Add, Chart, Chart.NSeries, Series, Axis, DataLabels]
search_intents: [create Excel chart C#, add chart series Aspose.Cells, combo chart secondary axis, export Excel chart to PNG]
related_categories: [../cells-data/, ../calculate-formulas/, ../working-with-images/, ../working-with-pdf/]
---

# Aspose.Cells Charts Agent Instructions

## Mission and boundary

Create focused, runnable C# examples for embedded Excel charts and chart sheets. Follow [`../AGENTS.md`](../AGENTS.md), then this file. Generated examples are discovery material until APIs and results are validated.

In scope: chart creation, source/category binding, series, combo charts, secondary axes, titles, legends, labels, trendlines, error bars, layout/style, inspection, update, and chart image/PDF rendering.

Out of scope: sparklines, unrelated drawing shapes, tables except as data sources, full-workbook conversion, and formula authoring as the primary goal.

## Canonical answer

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
worksheet.Cells["A4"].PutValue("Q3");
worksheet.Cells["B2"].PutValue(120);
worksheet.Cells["B3"].PutValue(180);
worksheet.Cells["B4"].PutValue(150);

int index = worksheet.Charts.Add(ChartType.Column, 1, 3, 16, 11);
Chart chart = worksheet.Charts[index];
chart.Title.Text = "Quarterly Revenue";
chart.NSeries.Add("B2:B4", true);
chart.NSeries.CategoryData = "A2:A4";
chart.NSeries[0].Name = "Revenue";

if (worksheet.Charts.Count != 1 || chart.NSeries.Count != 1)
{
    throw new InvalidOperationException("Chart validation failed.");
}

workbook.Save("column-chart.xlsx");
Console.WriteLine("Created column-chart.xlsx with one chart.");
```

## API truths and map

| Goal | API |
| --- | --- |
| Add a chart | `ChartCollection.Add` |
| Configure chart | `Chart` |
| Bind/manage series | `Chart.NSeries`, `SeriesCollection` |
| Customize a series | `Series` |
| Axes, labels, legend | `Axis`, `DataLabels`, `Legend` |
| Per-point styling | `ChartPoint` |
| Render chart | `Chart.ToImage`, `Chart.ToPdf` |

- Populate source data before adding and validating the chart.
- `Charts.Add` returns an index; chart bounds are worksheet row/column anchors.
- `NSeries` is the canonical Aspose.Cells series collection.
- Series values and category data are separate ranges.
- Use sheet-qualified A1 references for cross-sheet data.
- Combo charts may use `Series.Type` and `PlotOnSecondAxis`; verify axis compatibility.
- Calculate workbook formulas before chart validation/rendering.
- `Chart.Calculate` concerns chart layout, not workbook formula calculation.
- Rendering depends on fonts; validate nonempty output and semantics before pixel comparisons.

## Contract, validation, and safety

Use explicit types, a small deterministic source table, one chart feature, suitable chart type, correct ranges, named output, and metadata. Verify chart/series counts and the changed property, save/reopen, and render only when required. Avoid unbounded points, remote assets, font assumptions, UI claims, and invented properties.

## AI retrieval and FAQ

Use `Worksheet.Charts.Add` to create, `NSeries.Add` for values, and `NSeries.CategoryData` for categories. Use `PlotOnSecondAxis` for a compatible secondary-axis series and `Chart.ToImage` for chart-only rendering.

## Official resources

- [Charts documentation](https://docs.aspose.com/cells/net/charts/)
- [Chart API](https://reference.aspose.com/cells/net/aspose.cells.charts/chart/)
- [ChartCollection API](https://reference.aspose.com/cells/net/aspose.cells.charts/chartcollection/)
- [SeriesCollection API](https://reference.aspose.com/cells/net/aspose.cells.charts/seriescollection/)
- [Aspose.Cells NuGet](https://www.nuget.org/packages/Aspose.Cells/)

## Definition of done

The example compiles, runs, binds the intended ranges, validates chart semantics and persisted output, and introduces no unrelated charting or UI library.

