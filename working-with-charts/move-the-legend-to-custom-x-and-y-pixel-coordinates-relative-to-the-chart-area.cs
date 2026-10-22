// Title: Place a chart legend at exact pixel coordinates in an Aspose.Cells .NET column chart
// AI Prompts: Assign chart.Legend.X and chart.Legend.Y after converting 50 px and 30 px to points to position the legend relative to the chart area in C#. | Calculate points per pixel (72/96) and use those values to move the legend to a custom location in an Aspose.Cells workbook.
// Common Searches: asp.net aspose.cells set legend X and Y coordinates | c# how to move Excel chart legend by pixels using Aspose.Cells | pixel to point conversion for chart elements in Aspose.Cells | custom legend placement in column chart with Aspose.Cells | set legend position programmatically Aspose.Cells .NET
// Tags: chart legend pixel positioning Aspose.Cells | legend X Y property usage C# | convert pixel to point Aspose.Cells | custom legend location Excel chart .NET | column chart legend placement Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a workbook, adds sample data and a column chart, converts pixel offsets to points, sets the legend's X and Y properties to place it at the desired coordinates, and saves the file as ChartWithCustomLegend.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set legend to a predefined position (custom position not available in this version)
            chart.Legend.Position = LegendPositionType.Bottom;

            // Desired pixel coordinates relative to the chart area
            int pixelX = 50;
            int pixelY = 30;

            // Convert pixels to points (1 point = 1/72 inch, typical screen DPI = 96)
            double pointsPerPixel = 72.0 / 96.0; // 0.75
            chart.Legend.X = (int)(pixelX * pointsPerPixel);
            chart.Legend.Y = (int)(pixelY * pointsPerPixel);

            // Save the workbook
            string outputPath = "ChartWithCustomLegend.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
