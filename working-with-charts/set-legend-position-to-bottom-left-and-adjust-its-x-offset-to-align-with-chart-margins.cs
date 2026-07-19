// Title: Aspose.Cells for .NET – Set Chart Legend to Bottom‑Left and Align X Offset with Left Margin
// Description: This C# example creates a workbook, adds sample data, inserts a column chart, then moves the legend to the bottom of the plot area, aligns it with the left chart margin using XRatioToChart = 0.0, disables overlay, and saves the result as LegendBottomLeft.xlsx.
// Keywords: Aspose.Cells chart legend position | bottom left legend Aspose.Cells | XRatioToChart property C# | Legend.IsOverLay false | Aspose.Cells column chart example | C# Excel chart formatting
// Common Searches: how to place a chart legend at the bottom left in Aspose.Cells | align legend X offset with chart margin Aspose.Cells .NET | remove legend overlay from Excel chart using Aspose.Cells | set legend position and XRatioToChart in C#
// Developer Intent: Position a chart legend at the bottom‑left corner and align its horizontal offset with the left edge of the chart area in Aspose.Cells for .NET.
// Use Cases: Design financial dashboards where legends line up uniformly across multiple charts. | Generate reports that require non‑overlapping legends positioned consistently at the bottom left. | Automate Excel workbook creation with standardized chart layouts for corporate branding.
// AI Prompts: Show C# code to set a chart legend to bottom‑left and align its X offset using Aspose.Cells. | Explain the impact of Legend.IsOverLay and XRatioToChart when positioning legends in Aspose.Cells charts. | Provide a step‑by‑step guide to format chart legends for consistent layout in an Excel workbook with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendExample
{
    // This C# example creates a workbook, adds sample data, inserts a column chart, then moves the legend to the bottom of the plot area, aligns it with the left chart margin using XRatioToChart = 0.0, disables overlay, and saves the result as LegendBottomLeft.xlsx.
    class Program
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

            // Set legend to bottom‑left
            chart.Legend.Position = LegendPositionType.Bottom;   // bottom side of the plot area
            chart.Legend.XRatioToChart = 0.0;                    // align the legend with the left chart margin
            chart.Legend.IsOverLay = false;                     // ensure the legend does not overlap the plot area

            // Save the workbook
            workbook.Save("LegendBottomLeft.xlsx");
        }
    }
}
