// Title: Center a Chart Legend in Aspose.Cells for .NET without Overlap
// Description: Demonstrates how to create a column chart in a new workbook, disable legend overlay, set the legend to NotDocked, define its width and height as ratios of the chart, calculate X/Y ratios to place the legend at the plot‑area center, verify the settings with chart.Calculate(), and save the result as ChartWithCenteredLegend.xlsx.
// Keywords: Aspose.Cells legend center | C# chart legend position | Aspose.Cells NotDocked legend | prevent legend overlap Aspose.Cells | XRatioToChart YRatioToChart | column chart legend placement .NET | Aspose.Cells chart layout
// Common Searches: Aspose.Cells center legend in chart | how to prevent legend overlap in Aspose.Cells | set manual legend coordinates Aspose.Cells C# | legend.Position NotDocked Aspose.Cells example | center legend using XRatioToChart YRatioToChart
// Developer Intent: Place a chart legend at the exact center of the plot area while ensuring it does not cover any data series.
// Use Cases: Design sales dashboards where a balanced, centered legend improves readability. | Generate automated Excel reports with multiple charts that require consistent, non‑overlapping legend placement. | Programmatically validate visual layout of exported charts in data‑analysis pipelines.
// AI Prompts: Write C# code with Aspose.Cells to add a line chart and center its legend using NotDocked and ratio properties, ensuring no overlap. | Explain the effect of XRatioToChart and YRatioToChart on legend positioning in Aspose.Cells and show sample calculations. | Provide a step‑by‑step guide to verify legend placement after calling chart.Calculate() in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendCenterDemo
{
    // Demonstrates how to create a column chart in a new workbook, disable legend overlay, set the legend to NotDocked, define its width and height as ratios of the chart, calculate X/Y ratios to place the legend at the plot‑area center, verify the settings with chart.Calculate(), and save the result as ChartWithCenteredLegend.xlsx.
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
            chart.SetChartDataRange("A1:B4", true);

            // Access the legend
            Legend legend = chart.Legend;

            // Ensure the legend does not overlap the chart area
            legend.IsOverLay = false; // Show legend without overlapping the chart

            // Position the legend in the center of the plot area
            // Use NotDocked so we can set manual coordinates
            legend.Position = LegendPositionType.NotDocked;

            // Define legend size as a fraction of the chart area
            legend.WidthRatioToChart = 0.2;   // 20% of chart width
            legend.HeightRatioToChart = 0.1;  // 10% of chart height

            // Center the legend
            legend.XRatioToChart = 0.5 - legend.WidthRatioToChart / 2;
            legend.YRatioToChart = 0.5 - legend.HeightRatioToChart / 2;

            // Recalculate chart layout after manual positioning
            chart.Calculate();

            // Verify that the legend is set not to overlap
            Console.WriteLine("Legend IsOverLay (should be false): " + legend.IsOverLay);
            Console.WriteLine($"Legend Center Position - XRatio: {legend.XRatioToChart}, YRatio: {legend.YRatioToChart}");

            // Save the workbook
            workbook.Save("ChartWithCenteredLegend.xlsx");
        }
    }
}
