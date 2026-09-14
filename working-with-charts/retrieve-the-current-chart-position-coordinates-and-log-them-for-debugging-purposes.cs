// Title: How to retrieve and log an Aspose.Cells chart’s cell position, pixel size, and data point coordinates in C#
// AI Prompts: Write C# code that uses Aspose.Cells to read a chart’s UpperLeftRow, UpperLeftColumn, Width, and Height properties and output them to the console. | Show how to invoke Chart.Calculate() and then extract the ShapeXPx, ShapeYPx, ShapeX, and ShapeY values of a specific ChartPoint for debugging. | Demonstrate saving the workbook after logging chart placement and point coordinate information with Aspose.Cells.
// Common Searches: Aspose.Cells C# get chart upper left cell row and column | How to read pixel coordinates of a chart point using Aspose.Cells .NET | Debugging chart position and size in an Excel file with Aspose.Cells | Retrieve chart shape X Y ratio values Aspose.Cells C# example
// Tags: retrieve chart cell position Aspose.Cells | log chart pixel dimensions C# | calculate chart shape coordinates Aspose.Cells | debug excel chart placement Aspose.Cells | read chart point shape values .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartDebug
{
    // The example creates a workbook, adds sample data and a column chart, calls Chart.Calculate() to populate shape metrics, then logs the chart object's UpperLeftRow/UpperLeftColumn, pixel Width/Height, and the first series point's pixel (ShapeXPx/ShapeYPx) and ratio (ShapeX/ShapeY) coordinates before saving the file.
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

                // Add a column chart (rows 5‑20, columns 0‑8)
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Calculate the chart so that shape positions are populated
                chart.Calculate();

                // Retrieve the first point of the first series
                ChartPoint point = chart.NSeries[0].Points[0];

                // Log chart object's upper‑left cell position
                Console.WriteLine("Chart upper‑left cell: Row = " + chart.ChartObject.UpperLeftRow +
                                  ", Column = " + chart.ChartObject.UpperLeftColumn);

                // Log chart object's pixel size (optional, useful for debugging)
                Console.WriteLine("Chart size (pixels): Width = " + chart.ChartObject.Width +
                                  ", Height = " + chart.ChartObject.Height);

                // Log the point's position in pixels
                Console.WriteLine("Chart point position (pixels): X = " + point.ShapeXPx +
                                  ", Y = " + point.ShapeYPx);

                // Log the point's position in 1/4000 of chart width/height
                Console.WriteLine("Chart point position (ratio units): X = " + point.ShapeX +
                                  " (1/4000 of width), Y = " + point.ShapeY + " (1/4000 of height)");

                // Determine output file path and ensure directory exists
                string outputFile = Path.Combine(Directory.GetCurrentDirectory(), "ChartPositionDebug.xlsx");
                string outputDir = Path.GetDirectoryName(outputFile);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook (debug file)
                workbook.Save(outputFile);
                Console.WriteLine("Workbook saved to: " + outputFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            RetrieveChartPosition.Run();
        }
    }
}
