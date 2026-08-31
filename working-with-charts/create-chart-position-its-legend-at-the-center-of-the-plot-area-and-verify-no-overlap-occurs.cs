// Title: Create a column chart in Aspose.Cells for .NET and center its legend inside the plot area without overlap
// AI Prompts: Write C# code that uses Aspose.Cells to add a column chart, set the legend to NotDocked, and place it at the middle of the plot area by assigning XRatioToChart = 0.5 and YRatioToChart = 0.5. | Show how to enable legend overlay (IsOverLay = true) and recalculate the chart to ensure the centered legend does not cover the data series. | Demonstrate printing the legend’s Position, XRatioToChart, YRatioToChart, and IsOverLay values to the console and saving the workbook as an .xlsx file.
// Common Searches: Aspose.Cells .NET center chart legend inside plot area | C# set legend position NotDocked Aspose.Cells column chart | avoid legend overlapping data series in Aspose.Cells chart example | how to use XRatioToChart and YRatioToChart in Aspose.Cells | verify legend placement after centering in Aspose.Cells workbook
// Tags: Aspose.Cells column chart legend positioning | legend NotDocked mode Aspose.Cells | chart legend coordinate ratios Aspose.Cells | prevent legend covering data series Aspose.Cells | save workbook LegendCenteredChart.xlsx

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendCenterDemo
{
    // The example creates a new workbook, fills it with sample data, adds a column chart, configures the legend to NotDocked, positions it at the chart’s midpoint using XRatioToChart and YRatioToChart set to 0.5, enables overlay to avoid covering the plot area, recalculates the chart, outputs legend properties to the console, and saves the file as LegendCenteredChart.xlsx.
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

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Position the legend at the center of the plot area
            // Use NotDocked so we can set manual coordinates
            chart.Legend.Position = LegendPositionType.NotDocked;
            // Center horizontally and vertically (ratio values are between 0 and 1)
            chart.Legend.XRatioToChart = 0.5; // center X
            chart.Legend.YRatioToChart = 0.5; // center Y
            // Ensure the legend does not overlap the chart plot area
            chart.Legend.IsOverLay = true;

            // Recalculate the chart to apply the positioning
            chart.Calculate();

            // Simple verification: output the legend settings to the console
            Console.WriteLine($"Legend Position: {chart.Legend.Position}");
            Console.WriteLine($"Legend X Ratio: {chart.Legend.XRatioToChart}");
            Console.WriteLine($"Legend Y Ratio: {chart.Legend.YRatioToChart}");
            Console.WriteLine($"Legend IsOverLay: {chart.Legend.IsOverLay}");

            // Save the workbook
            workbook.Save("LegendCenteredChart.xlsx");
        }
    }
}
