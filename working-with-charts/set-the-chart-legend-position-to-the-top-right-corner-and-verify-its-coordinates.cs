// Title: Position a chart legend in the top‑right corner and read its ratio‑to‑chart and pixel coordinates with Aspose.Cells for .NET
// AI Prompts: Place the legend of a column chart at the top‑right corner, then output its XRatioToChart, YRatioToChart, XPixel, and YPixel values. | After setting Legend.Position = Corner, call chart.Calculate() and retrieve the legend's relative and pixel coordinates using Aspose.Cells. | Adjust the example to compare the retrieved legend coordinates with the expected top‑right values for verification.
// Common Searches: how to set Aspose.Cells chart legend to top right corner in C# | retrieve legend XRatioToChart and YRatioToChart values with Aspose.Cells .NET | get pixel coordinates of a chart legend using Aspose.Cells for .NET | verify legend placement after positioning in an Aspose.Cells chart | Aspose.Cells calculate chart to update legend position coordinates
// Tags: Aspose.Cells chart legend corner positioning | Legend.XRatioToChart retrieval .NET | chart legend pixel coordinates Aspose.Cells | column chart legend top right placement | calculate chart to update legend position Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendDemo
{
    // Creates a workbook, adds a column chart, positions its legend at the top‑right corner, forces chart calculation, and prints the legend's ratio‑to‑chart and pixel coordinates before saving the file.
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
            chart.SetChartDataRange("A1:B4", true);

            // Position the legend at the top‑right corner of the plot area
            chart.Legend.Position = LegendPositionType.Corner;

            // Force calculation so that position properties are up‑to‑date
            chart.Calculate();

            // Retrieve legend coordinates (ratio to chart area)
            double legendXRatio = chart.Legend.XRatioToChart;
            double legendYRatio = chart.Legend.YRatioToChart;

            // Also retrieve pixel coordinates for verification
            int legendXPixel = chart.Legend.XPixel;
            int legendYPixel = chart.Legend.YPixel;

            // Output the coordinates to the console
            Console.WriteLine($"Legend Position: {chart.Legend.Position}");
            Console.WriteLine($"Legend X Ratio to Chart: {legendXRatio}");
            Console.WriteLine($"Legend Y Ratio to Chart: {legendYRatio}");
            Console.WriteLine($"Legend X Pixel: {legendXPixel}");
            Console.WriteLine($"Legend Y Pixel: {legendYPixel}");

            // Save the workbook
            workbook.Save("LegendTopRightCorner.xlsx");
        }
    }
}
