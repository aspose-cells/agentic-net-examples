// Title: Aspose.Cells for .NET – Retrieve Chart Cell Bounds and Data‑Point Pixel Coordinates
// Description: This example creates a workbook, adds a column chart, populates it with sample data, calls Calculate to finalize layout, then reads the chart object's UpperLeftRow, UpperLeftColumn, LowerRightRow and LowerRightColumn. It also logs the first data point's pixel coordinates (ShapeXPx, ShapeYPx) and relative positions (ShapeX, ShapeY) before saving the file.
// Keywords: Aspose.Cells chart position | ChartObject UpperLeftRow | ChartObject LowerRightColumn | ShapeXPx ShapeYPx | retrieve chart bounds .NET | Excel chart pixel coordinates | C# Aspose.Cells chart debugging
// Common Searches: how to get chart cell coordinates with Aspose.Cells | Aspose.Cells retrieve chart upper left row column | pixel location of chart data point Aspose .NET | chart object bounds Aspose.Cells example | log chart point ShapeX ShapeY Aspose
// Developer Intent: Obtain a chart's worksheet cell boundaries and the exact pixel location of its data points for debugging or layout calculations.
// Use Cases: Debug overlapping charts by logging their cell bounds. | Align multiple charts programmatically using UpperLeft/LowerRight coordinates. | Create custom annotations positioned at specific data‑point pixels. | Convert ShapeX/ShapeY (1/4000 of chart size) to absolute pixels for precise rendering.
// AI Prompts: Show C# code that reads a chart's UpperLeftRow, UpperLeftColumn, LowerRightRow, and LowerRightColumn using Aspose.Cells. | Provide an example that iterates all series points in an Aspose.Cells chart and prints ShapeX, ShapeY, ShapeXPx, and ShapeYPx. | Explain how to convert ShapeX/ShapeY values to actual pixel coordinates for accurate overlay placement.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This example creates a workbook, adds a column chart, populates it with sample data, calls Calculate to finalize layout, then reads the chart object's UpperLeftRow, UpperLeftColumn, LowerRightRow and LowerRightColumn. It also logs the first data point's pixel coordinates (ShapeXPx, ShapeYPx) and relative positions (ShapeX, ShapeY) before saving the file.
    public class RetrieveChartPositionDemo
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

                // Add a column chart (topRow, leftColumn, bottomRow, rightColumn)
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Calculate the chart so that shape positions are populated
                chart.Calculate();

                // Retrieve chart object's upper‑left and lower‑right cell coordinates
                Console.WriteLine("Chart upper‑left row: " + chart.ChartObject.UpperLeftRow);
                Console.WriteLine("Chart upper‑left column: " + chart.ChartObject.UpperLeftColumn);
                Console.WriteLine("Chart lower‑right row: " + chart.ChartObject.LowerRightRow);
                Console.WriteLine("Chart lower‑right column: " + chart.ChartObject.LowerRightColumn);

                // Retrieve the first point of the first series
                ChartPoint point = chart.NSeries[0].Points[0];

                // Log the point's position in pixels
                Console.WriteLine("Chart point ShapeXPx (X in pixels): " + point.ShapeXPx);
                Console.WriteLine("Chart point ShapeYPx (Y in pixels): " + point.ShapeYPx);

                // Also log the relative positions (1/4000 of chart width/height)
                Console.WriteLine("Chart point ShapeX (1/4000 of width): " + point.ShapeX);
                Console.WriteLine("Chart point ShapeY (1/4000 of height): " + point.ShapeY);

                // Save the workbook (optional, just to complete lifecycle)
                string outputPath = "RetrieveChartPositionDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            RetrieveChartPositionDemo.Run();
        }
    }
}
