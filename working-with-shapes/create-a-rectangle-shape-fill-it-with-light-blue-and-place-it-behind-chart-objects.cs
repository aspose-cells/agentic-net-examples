// Title: C# Example – Insert a Light‑Blue Rectangle Behind a Chart Using Aspose.Cells
// Description: This Aspose.Cells for .NET sample creates a workbook, adds sample data and a column chart, then inserts a rectangle shape inside the chart with AddShapeInChartByScale, fills it with LightBlue, sets its ZOrderPosition to send it to the back, and saves the file as RectangleBehindChart.xlsx.
// Keywords: Aspose.Cells | C# chart shape | AddShapeInChartByScale | rectangle behind chart | ZOrderPosition | Excel chart background | light blue fill | .NET | shape Z‑order | chart watermark | Aspose.Cells example | GitHub Aspose.Cells | Excel shape layering
// Common Searches: Aspose.Cells add rectangle behind chart C# | How to set shape Z‑order in Aspose.Cells chart | Add background shape to Excel chart using Aspose.Cells | C# example for AddShapeInChartByScale | Place shape behind chart objects Aspose.Cells
// Developer Intent: Add a light‑blue rectangle shape behind a chart in an Excel workbook with Aspose.Cells for .NET.
// Use Cases: Create a colored background for a chart while keeping data series visible. | Add a subtle watermark behind chart elements for branding or copyright notices. | Design custom chart layouts by layering solid‑fill shapes beneath chart series.
// AI Prompts: Generate C# code that inserts a rectangle shape behind a chart using Aspose.Cells and applies a LightBlue solid fill. | Show how to adjust a shape's ZOrderPosition in an Aspose.Cells chart so the shape appears behind all chart objects. | Provide an Aspose.Cells example that uses AddShapeInChartByScale to position a shape relative to chart dimensions and set its fill color.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // This Aspose.Cells for .NET sample creates a workbook, adds sample data and a column chart, then inserts a rectangle shape inside the chart with AddShapeInChartByScale, fills it with LightBlue, sets its ZOrderPosition to send it to the back, and saves the file as RectangleBehindChart.xlsx.
    public class RectangleBehindChartDemo
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

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 1, 20, 10);
                Chart chart = worksheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Add a rectangle shape inside the chart using percent scale coordinates
                // left=10%, top=10%, width=80%, height=80% of the chart area
                Shape rectangle = chart.Shapes.AddShapeInChartByScale(
                    MsoDrawingType.Rectangle,
                    PlacementType.Move,
                    0.1,   // left (10%)
                    0.1,   // top (10%)
                    0.8,   // width (80%)
                    0.8);  // height (80%)

                // Fill the rectangle with light blue color
                rectangle.Fill.SolidFill.Color = Color.LightBlue;

                // Place the rectangle behind all chart objects
                rectangle.ZOrderPosition = 0; // lower Z-order means back

                // Save the workbook
                workbook.Save("RectangleBehindChart.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            RectangleBehindChartDemo.Run();
        }
    }
}
