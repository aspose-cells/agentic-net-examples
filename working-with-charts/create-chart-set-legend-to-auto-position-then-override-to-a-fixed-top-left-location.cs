// Title: Aspose.Cells for .NET – Create a Column Chart and Position the Legend at a Fixed Top‑Left Location
// Description: C# example that builds a workbook, adds sample data, inserts a column chart, sets the legend to automatic placement, then switches to NotDocked and uses XRatioToChart/YRatioToChart (with width and height ratios) to lock the legend at the chart’s top‑left corner before saving.
// Keywords: Aspose.Cells | C# chart legend | NotDocked legend | XRatioToChart | YRatioToChart | legend size ratios | automatic legend positioning | column chart .NET | Aspose.Cells example
// Common Searches: Aspose.Cells set legend position C# | How to move chart legend to top left in Aspose.Cells | NotDocked legend XRatioToChart Aspose.Cells | Override automatic legend placement Aspose.Cells .NET | Set legend width and height ratios Aspose.Cells
// Developer Intent: Create a column chart and programmatically place its legend at a specific top‑left coordinate after applying automatic positioning.
// Use Cases: Standardize legend placement in sales dashboards for consistent visual layout. | Generate printable reports where the legend must stay outside the data area. | Automate workbook creation with precise legend sizing for multi‑sheet analytics.
// AI Prompts: Show C# code to set a chart legend to NotDocked and define XRatioToChart/YRatioToChart in Aspose.Cells. | Provide an Aspose.Cells example that first uses automatic legend positioning then moves the legend to a custom top‑left spot with width and height ratios. | Explain the difference between automatic legend docking and NotDocked with ratio properties in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// C# example that builds a workbook, adds sample data, inserts a column chart, sets the legend to automatic placement, then switches to NotDocked and uses XRatioToChart/YRatioToChart (with width and height ratios) to lock the legend at the chart’s top‑left corner before saving.
class ChartLegendPositionDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Access the legend of the chart
        Legend legend = chart.Legend;

        // 1. Set legend to automatic positioning
        legend.SetPositionAuto();

        // 2. Override to a fixed top‑left location
        //    Use NotDocked so that X/Y coordinates are respected
        legend.Position = LegendPositionType.NotDocked;

        // Position the legend at the top‑left corner of the chart area
        // X/Y ratios are fractions of the chart area (0.0 = left/top, 1.0 = right/bottom)
        legend.XRatioToChart = 0.0;   // left edge
        legend.YRatioToChart = 0.0;   // top edge

        // Optionally set size of the legend (20% width, 20% height of chart area)
        legend.WidthRatioToChart = 0.2;
        legend.HeightRatioToChart = 0.2;

        // Save the workbook
        workbook.Save("ChartLegendPositionDemo.xlsx");
    }
}
