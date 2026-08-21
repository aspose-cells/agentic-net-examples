// Title: Get Chart Point Coordinates (Pixel & Relative) with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds a column chart, runs chart.Calculate() and reads the first point's ShapeXPx, ShapeYPx, ShapeX and ShapeY values, then logs them for debugging before saving the file.
// Keywords: Aspose.Cells chart point coordinates | ChartPoint ShapeXPx ShapeYPx | Aspose.Cells chart.Calculate | C# retrieve chart shape position | Aspose.Cells debugging chart layout | relative chart point units 1/4000 | Aspose.Cells get pixel location
// Common Searches: how to read pixel position of a chart point in Aspose.Cells | Aspose.Cells chart.Calculate to obtain shape coordinates | retrieve ShapeX and ShapeY values from a chart series C# | log chart point location for debugging Aspose.Cells | convert chart point relative units to percentage Aspose.Cells
// Developer Intent: Extract X/Y coordinates of a chart point and output them for troubleshooting.
// Use Cases: Verify exact placement of data points when programmatically adjusting chart size. | Compare relative ShapeX/ShapeY values to align custom graphics over a chart. | Capture point coordinates for export to external reporting tools.
// AI Prompts: Show C# code that uses Aspose.Cells to read ShapeXPx and ShapeYPx of the first point after calling chart.Calculate(). | Explain how to translate ShapeX and ShapeY (1/4000 of chart size) into percentage values of the chart area. | Suggest ways to write chart point coordinates to a log file instead of the console in a .NET application.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartDebug
{
    // Creates a workbook, adds a column chart, runs chart.Calculate() and reads the first point's ShapeXPx, ShapeYPx, ShapeX and ShapeY values, then logs them for debugging before saving the file.
    public class RetrieveChartPosition
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(30);

                // Add a column chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Calculate the chart so that shape position properties are populated
                chart.Calculate();

                // Retrieve the first point of the first series
                ChartPoint point = chart.NSeries[0].Points[0];

                // Log the current X and Y coordinates in pixels
                Console.WriteLine("Chart Point Position (pixels):");
                Console.WriteLine($"  ShapeXPx = {point.ShapeXPx}");
                Console.WriteLine($"  ShapeYPx = {point.ShapeYPx}");

                // Also log the relative coordinates (1/4000 of chart size)
                Console.WriteLine("Chart Point Position (relative units):");
                Console.WriteLine($"  ShapeX = {point.ShapeX}");
                Console.WriteLine($"  ShapeY = {point.ShapeY}");

                // Save the workbook (optional, just to keep the example complete)
                workbook.Save("ChartPositionDebug.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            RetrieveChartPosition.Run();
        }
    }
}
