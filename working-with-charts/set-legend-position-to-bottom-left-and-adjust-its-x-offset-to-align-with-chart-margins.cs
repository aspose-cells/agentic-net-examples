// Title: Set the chart legend to the bottom‑left corner and align its X offset with the chart margins in Aspose.Cells for .NET
// AI Prompts: Create a column chart with Aspose.Cells and move the legend to the bottom left using Legend.Position and XRatioToChart. | Adjust an Aspose.Cells chart so the legend sits at the bottom edge without overlapping the plot area by setting IsOverLay to false. | Generate an Excel workbook where the legend’s X offset is set to zero to align it with the left margin of the chart.
// Common Searches: aspocells how to place chart legend at bottom left in C# | set legend XRatioToChart to align with chart margins Aspose.Cells | prevent legend overlay when positioning legend below chart Aspose.Cells .NET | C# Aspose.Cells legend position bottom left offset example
// Tags: Aspose.Cells legend bottom positioning | Aspose.Cells XRatioToChart alignment | Aspose.Cells column chart legend placement | Aspose.Cells disable legend overlay | Aspose.Cells chart margin alignment

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, inserts a column chart, sets the legend position to Bottom, aligns it to the left edge of the chart by setting XRatioToChart to 0, disables overlay to avoid overlap, and saves the file as LegendBottomLeft.xlsx.
class SetLegendBottomLeft
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

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Access the legend and set it to the bottom of the plot area
        Legend legend = chart.Legend;
        legend.Position = LegendPositionType.Bottom;

        // Align the legend with the left margin of the chart.
        // When the legend is positioned at the bottom, the X coordinate is respected.
        // Using XRatioToChart = 0 places it at the very left edge of the chart area.
        legend.XRatioToChart = 0.0;

        // Optional: ensure the legend does not overlap the chart
        legend.IsOverLay = false;

        // Save the workbook
        workbook.Save("LegendBottomLeft.xlsx");
    }
}
