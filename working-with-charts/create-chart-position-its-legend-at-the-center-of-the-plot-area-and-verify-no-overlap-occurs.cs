// Title: Center Chart Legend Inside Plot Area with Aspose.Cells for .NET
// Description: Creates a workbook, adds a column chart, sets the legend to NotDocked with overlay, defines a fixed size, computes the plot‑area ratios, positions the legend at the plot‑area centre, recalculates the layout, checks that the legend stays fully inside the plot bounds, and saves the file.
// Keywords: Aspose.Cells | C# chart legend | center legend | plot area | NotDocked legend | legend overlay | chart layout | prevent legend overlap | Aspose.Cells example
// Common Searches: Aspose.Cells center legend plot area | C# position chart legend inside plot area | how to prevent legend overlap in Aspose.Cells chart | set legend NotDocked and overlay Aspose.Cells | validate legend bounds Aspose.Cells
// Developer Intent: Programmatically place a chart legend at the exact centre of the plot area and confirm it does not exceed the plot boundaries.
// Use Cases: Generate a column chart and manually centre a fixed‑size legend for a tidy visual layout. | Automatically verify legend coordinates before saving to avoid clipping or overlap. | Adapt legend size and position dynamically based on plot‑area dimensions for responsive chart designs.
// AI Prompts: Write C# code using Aspose.Cells that adds a chart, sets the legend to NotDocked with overlay, defines width/height ratios, and centers it inside the plot area. | Create a method that returns true if an Aspose.Cells chart legend is completely contained within the plot area. | Show an example that moves the legend to the centre of the plot area for any chart type and logs whether the placement is inside the bounds.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendCenterDemo
{
    // Creates a workbook, adds a column chart, sets the legend to NotDocked with overlay, defines a fixed size, computes the plot‑area ratios, positions the legend at the plot‑area centre, recalculates the layout, checks that the legend stays fully inside the plot bounds, and saves the file.
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
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the legend
            Legend legend = chart.Legend;

            // Place legend inside the plot area (not docked) and allow overlay
            legend.Position = LegendPositionType.NotDocked;
            legend.IsOverLay = true; // Enable overlay so we can position it manually

            // Retrieve plot area dimensions (ratios to chart)
            PlotArea plotArea = chart.PlotArea;
            double plotX = plotArea.XRatioToChart;
            double plotY = plotArea.YRatioToChart;
            double plotWidth = plotArea.WidthRatioToChart;
            double plotHeight = plotArea.HeightRatioToChart;

            // Ensure legend has a defined size (optional, otherwise automatic)
            legend.IsAutomaticSize = false;
            legend.WidthRatioToChart = 0.2;   // 20% of chart width
            legend.HeightRatioToChart = 0.1;  // 10% of chart height

            // Center the legend within the plot area
            legend.XRatioToChart = plotX + (plotWidth - legend.WidthRatioToChart) / 2;
            legend.YRatioToChart = plotY + (plotHeight - legend.HeightRatioToChart) / 2;

            // Recalculate chart layout to apply changes
            chart.Calculate();

            // Verify that legend lies completely inside the plot area (no overlap outside)
            bool isInsidePlotArea =
                legend.XRatioToChart >= plotX &&
                legend.YRatioToChart >= plotY &&
                (legend.XRatioToChart + legend.WidthRatioToChart) <= (plotX + plotWidth) &&
                (legend.YRatioToChart + legend.HeightRatioToChart) <= (plotY + plotHeight);

            Console.WriteLine(isInsidePlotArea
                ? "Legend is centered within the plot area with no external overlap."
                : "Legend placement exceeds plot area bounds.");

            // Save the workbook
            workbook.Save("LegendCentered.xlsx");
        }
    }
}
