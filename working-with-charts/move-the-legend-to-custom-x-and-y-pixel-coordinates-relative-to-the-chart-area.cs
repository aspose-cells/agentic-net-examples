// Title: Position an Excel chart legend at specific pixel coordinates using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set chart legend X and Y pixel offset | how to move Excel chart legend to exact pixel location with Aspose.Cells | convert pixel values to points for legend placement in Aspose.Cells chart | custom legend positioning relative to chart area in C# Aspose.Cells | adjust legend coordinates after setting bottom position Aspose.Cells
// Tags: chart legend pixel positioning Aspose.Cells | legend X Y offset points conversion | custom legend placement Excel chart C# | Aspose.Cells chart layout adjustments | pixel to point conversion for Excel elements

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a workbook, adds sample data, inserts a column chart, sets the legend to the Bottom position, converts desired pixel offsets to points, applies the X and Y coordinates to the legend, and saves the file as an Excel workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
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

            // Add a column chart to the worksheet (firstRow, firstColumn, totalRows, totalColumns)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // ---- Move legend to custom X and Y pixel coordinates ----
            // Use a standard position (e.g., Bottom) and then adjust X/Y manually
            chart.Legend.Position = LegendPositionType.Bottom;

            // Desired pixel coordinates relative to the chart area
            int pixelX = 150; // horizontal offset in pixels
            int pixelY = 30;  // vertical offset in pixels

            // Aspose.Cells uses points for positioning (1 point = 1/72 inch).
            // At a typical screen DPI of 96, 1 pixel = 0.75 point.
            const double pointsPerPixel = 0.75;

            // Apply the converted coordinates to the legend (cast to int because X/Y are int)
            chart.Legend.X = (int)(pixelX * pointsPerPixel);
            chart.Legend.Y = (int)(pixelY * pointsPerPixel);
            // ---------------------------------------------------------

            // Define output file path
            string outputPath = "ChartWithCustomLegend.xlsx";

            // Ensure the directory exists (handle case where outputPath has no directory part)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
