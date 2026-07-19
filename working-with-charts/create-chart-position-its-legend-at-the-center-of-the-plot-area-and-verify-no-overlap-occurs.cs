// Title: Center Chart Legend in Plot Area and Prevent Overlap with Aspose.Cells for C#
// Description: Demonstrates how to create a column chart with Aspose.Cells, make the legend free‑positioned, size it using ratio properties, calculate the plot area, center the legend via XRatioToChart/YRatioToChart, disable overlay, verify non‑overlap with a bounding‑box check, and save the workbook.
// Keywords: Aspose.Cells | C# | .NET | chart legend | center legend | plot area | NotDocked | XRatioToChart | YRatioToChart | prevent legend overlap | column chart | workbook export
// Common Searches: Aspose.Cells center legend in plot area C# | How to prevent chart legend overlap using Aspose.Cells | Set legend NotDocked and size ratios Aspose.Cells .NET | Calculate plot area and position legend Aspose.Cells | Verify legend bounds against plot area Aspose.Cells
// Developer Intent: Place a chart legend at the plot‑area center and ensure it does not cover the data series.
// Use Cases: Improve readability of financial column charts by centering the legend without obscuring bars. | Automate validation of legend placement in generated dashboards before distribution. | Dynamically adjust legend size and position for various chart types while guaranteeing no overlap.
// AI Prompts: Write C# code with Aspose.Cells that centers a chart legend inside the plot area and checks for overlap. | Explain how XRatioToChart and YRatioToChart are used to position a legend relative to the plot area in Aspose.Cells. | Suggest alternative methods to avoid legend overlap without using the IsOverLay property.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendCenterDemo
{
    // Demonstrates how to create a column chart with Aspose.Cells, make the legend free‑positioned, size it using ratio properties, calculate the plot area, center the legend via XRatioToChart/YRatioToChart, disable overlay, verify non‑overlap with a bounding‑box check, and save the workbook.
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

            // Access the legend
            Legend legend = chart.Legend;

            // Make the legend free‑positioned (not docked)
            legend.Position = LegendPositionType.NotDocked;

            // Set a reasonable size for the legend (as a fraction of the chart area)
            legend.WidthRatioToChart = 0.3;   // 30% of chart width
            legend.HeightRatioToChart = 0.2;  // 20% of chart height

            // Center the legend within the plot area
            // Plot area defaults to the whole chart area; after Calculate it will be accurate
            chart.Calculate(); // Ensure plot area dimensions are up‑to‑date

            PlotArea plot = chart.PlotArea;
            // Center X
            legend.XRatioToChart = plot.XRatioToChart + (plot.WidthRatioToChart - legend.WidthRatioToChart) / 2;
            // Center Y
            legend.YRatioToChart = plot.YRatioToChart + (plot.HeightRatioToChart - legend.HeightRatioToChart) / 2;

            // Ensure the legend does NOT overlap the chart plot area
            // Setting IsOverLay to false forces the legend to be placed outside if overlap would occur
            legend.IsOverLay = false;

            // Re‑calculate after positioning
            chart.Calculate();

            // Verify overlap (simple bounding‑box check)
            bool overlap =
                legend.XRatioToChart < plot.XRatioToChart + plot.WidthRatioToChart &&
                legend.XRatioToChart + legend.WidthRatioToChart > plot.XRatioToChart &&
                legend.YRatioToChart < plot.YRatioToChart + plot.HeightRatioToChart &&
                legend.YRatioToChart + legend.HeightRatioToChart > plot.YRatioToChart;

            Console.WriteLine(overlap
                ? "Warning: Legend overlaps the plot area."
                : "Success: Legend is centered and does not overlap the plot area.");

            // Save the workbook
            workbook.Save("LegendCenteredNoOverlap.xlsx");
        }
    }
}
