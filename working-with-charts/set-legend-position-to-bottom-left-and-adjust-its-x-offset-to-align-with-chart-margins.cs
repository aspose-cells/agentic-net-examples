// Title: Aspose.Cells for .NET – Position Chart Legend Bottom‑Left and Align with Chart Margins (C#)
// Description: Creates a workbook, adds sample data, inserts a column chart, then sets the chart's Legend.Position to Bottom and adjusts Legend.X to 0 so the legend aligns with the left edge of the chart area before saving the file as an Excel workbook.
// Keywords: Aspose.Cells legend bottom left | C# chart legend offset | Aspose.Cells set legend position | Excel chart legend alignment | Aspose.Cells Chart.Legend.X | .NET Excel chart formatting
// Common Searches: how to move chart legend to bottom left in Aspose.Cells | Aspose.Cells legend X property example | C# set chart legend position bottom Aspose.Cells | align Excel chart legend with left margin using Aspose
// Developer Intent: Place a chart legend at the bottom of the plot area and shift it horizontally so it lines up with the chart’s left margin.
// Use Cases: Automated financial reports where the legend must sit at the bottom left to preserve page layout. | Generating dashboards that require consistent legend placement across multiple charts. | Preparing printable spreadsheets where the legend should not overlap data series or worksheet content.
// AI Prompts: Show a C# snippet that sets a chart legend to Bottom and aligns it with the left edge using Aspose.Cells. | Explain the effect of the Legend.X property in Aspose.Cells and how to use it for precise legend positioning. | Provide step‑by‑step instructions to adjust a chart legend’s horizontal offset for Excel files created with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample data, inserts a column chart, then sets the chart's Legend.Position to Bottom and adjusts Legend.X to 0 so the legend aligns with the left edge of the chart area before saving the file as an Excel workbook.
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

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Access the legend and position it at the bottom
        Legend legend = chart.Legend;
        legend.Position = LegendPositionType.Bottom;   // Bottom side of the plot area

        // Align the legend with the left margin of the chart.
        // Since the legend is docked at the bottom, the X property (or XPixel) controls horizontal offset.
        // Setting X to 0 aligns it with the left edge of the chart area.
        legend.X = 0;   // offset in units of 1/4000 of the chart area (obsolete but functional)

        // Save the workbook
        workbook.Save("LegendBottomLeft.xlsx");
    }
}
