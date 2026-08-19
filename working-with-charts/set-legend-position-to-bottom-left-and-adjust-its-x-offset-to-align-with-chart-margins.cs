// Title: Aspose.Cells for .NET – Position Chart Legend at Bottom‑Left and Align X Offset
// Description: Demonstrates how to create a workbook, add a column chart, and configure the legend to sit at the bottom‑left of the plot area. The example sets Legend.Position, uses XRatioToChart = 0.0 to line up with the left margin, and disables overlay to keep the chart area clear, then saves the file as an XLSX document.
// Keywords: Aspose.Cells legend position | chart legend bottom left C# | XRatioToChart Aspose.Cells | disable legend overlay | .NET chart formatting | Aspose.Cells chart customization
// Common Searches: Aspose.Cells set legend to bottom left | C# align chart legend with left margin | how to use Legend.XRatioToChart in Aspose.Cells | remove legend overlay from chart Aspose.Cells | position chart legend bottom side .NET
// Developer Intent: Place a chart legend at the bottom‑left of the plot area and align its left edge with the chart margin without overlapping the data series.
// Use Cases: Designing financial dashboards where the legend must stay out of the plot area and align with the left edge for a tidy layout. | Generating automated reports that follow corporate branding rules requiring legends at a fixed bottom‑left position. | Creating reusable chart templates in .NET applications where legend placement and offset are programmatically controlled.
// AI Prompts: Generate C# code using Aspose.Cells to set a chart legend at the bottom‑left and align its X offset with the chart margin. | Show how to adjust Legend.XRatioToChart and disable overlay for a column chart in Aspose.Cells. | Provide an example that moves an existing chart legend to the bottom side and ensures it does not overlap the plot area.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a workbook, add a column chart, and configure the legend to sit at the bottom‑left of the plot area. The example sets Legend.Position, uses XRatioToChart = 0.0 to line up with the left margin, and disables overlay to keep the chart area clear, then saves the file as an XLSX document.
class SetLegendBottomLeft
{
    static void Main()
    {
        // Create a new workbook
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

        // Configure the legend: position at bottom left
        Legend legend = chart.Legend;
        legend.Position = LegendPositionType.Bottom;   // Bottom side of the plot area
        legend.XRatioToChart = 0.0;                    // Align left edge with chart margin
        legend.IsOverLay = false;                      // Ensure legend does not overlap the plot area

        // Save the workbook
        workbook.Save("LegendBottomLeft.xlsx");
    }
}
