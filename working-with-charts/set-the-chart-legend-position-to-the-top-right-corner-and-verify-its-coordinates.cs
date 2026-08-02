// Title: C# – Set Aspose.Cells Chart Legend to Top‑Right Corner and Get Its Position Ratios & Pixels
// Description: Demonstrates how to create a workbook with a column chart, place the legend in the top‑right corner using Aspose.Cells for .NET, disable overlay, recalculate the chart, and read the legend's X/Y ratios (relative to the chart) and absolute pixel coordinates before saving the file.
// Keywords: Aspose.Cells chart legend position | C# set legend top right | LegendPositionType.Corner | XRatioToChart | YRatioToChart | legend pixel coordinates | .NET Excel chart legend | Aspose.Cells legend overlay | retrieve legend coordinates | Aspose.Cells example
// Common Searches: Aspose.Cells set chart legend top right | How to get legend XRatioToChart in C# | Legend pixel location Aspose.Cells | C# chart legend corner position Aspose | Retrieve legend coordinates from Excel chart
// Developer Intent: Place a chart legend in the top‑right corner of an Aspose.Cells chart and read its relative ratios and absolute pixel coordinates.
// Use Cases: Ensure legends do not overlap data series by positioning them in the corner and confirming exact placement. | Dynamically adjust legend location based on chart size using XRatioToChart and YRatioToChart values. | Create standardized Excel reports where every chart’s legend appears consistently in the top‑right corner.
// AI Prompts: Generate C# code with Aspose.Cells that sets a chart legend to the top‑right corner and prints its X/Y ratios and pixel coordinates. | Explain the effect of Legend.Position = LegendPositionType.Corner on XRatioToChart and YRatioToChart in Aspose.Cells. | Write a validation method in C# that checks whether the legend lies within the chart bounds after applying the Corner position.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendPositionDemo
{
    // Demonstrates how to create a workbook with a column chart, place the legend in the top‑right corner using Aspose.Cells for .NET, disable overlay, recalculate the chart, and read the legend's X/Y ratios (relative to the chart) and absolute pixel coordinates before saving the file.
    class Program
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
            chart.SetChartDataRange("A1:B4", true);

            // Set the legend position to the top‑right corner of the plot area
            chart.Legend.Position = LegendPositionType.Corner;

            // Ensure the legend does not overlap the chart (optional)
            chart.Legend.IsOverLay = false;

            // Recalculate the chart so that layout properties (e.g., X/Y ratios) are updated
            chart.Calculate();

            // Retrieve and display the legend's position ratios and pixel coordinates
            double legendXRatio = chart.Legend.XRatioToChart; // 0‑1 fraction of chart width
            double legendYRatio = chart.Legend.YRatioToChart; // 0‑1 fraction of chart height
            int legendXPixel = chart.Legend.XPixel;           // X position in pixels
            int legendYPixel = chart.Legend.YPixel;           // Y position in pixels

            Console.WriteLine($"Legend Position Ratios -> X: {legendXRatio:F3}, Y: {legendYRatio:F3}");
            Console.WriteLine($"Legend Position Pixels  -> X: {legendXPixel}, Y: {legendYPixel}");

            // Save the workbook
            workbook.Save("LegendTopRightCorner.xlsx");
        }
    }
}
