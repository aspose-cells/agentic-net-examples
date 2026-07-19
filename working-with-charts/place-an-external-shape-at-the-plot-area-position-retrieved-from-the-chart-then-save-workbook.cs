// Title: Overlay a Rectangle on a Chart Plot Area with Aspose.Cells for .NET and Save the Workbook
// Description: Creates a new workbook, populates cells A1:B4, adds a column chart, calculates it, then inserts a rectangle shape that exactly matches the chart's plot area using AddShapeInChartByScale (0‑100% coordinates). The shape is styled with a light‑red fill and red border before the workbook is saved as ChartWithExternalShape.xlsx.
// Keywords: Aspose.Cells | C# chart shape | AddShapeInChartByScale | plot area overlay | external shape Excel chart | rectangle shape Aspose.Cells | chart styling .NET | save workbook with chart | Aspose.Cells example
// Common Searches: Aspose.Cells add shape to chart plot area | C# overlay rectangle on Excel chart using Aspose | AddShapeInChartByScale tutorial | how to match shape size to chart plot area Aspose.Cells | save workbook after inserting chart shape
// Developer Intent: Insert a rectangle that precisely covers a chart’s plot area and persist the modification in the Excel file.
// Use Cases: Emphasize the data region of a chart with a colored overlay in automated reports. | Provide a custom background for a chart by programmatically adding a shape that fits the plot area. | Prepare an Excel file for downstream processing where a shape marks the chart area for visual cues.
// AI Prompts: Generate C# code with Aspose.Cells that adds a rectangle shape aligned to a chart’s plot area using scale coordinates and saves the workbook. | Show how to retrieve a chart’s plot area dimensions and place an external shape or image over it in Aspose.Cells for .NET. | Provide an example that styles a shape added to a chart (fill color, border) and persists the changes in an Excel file.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartShapeExample
{
    // Creates a new workbook, populates cells A1:B4, adds a column chart, calculates it, then inserts a rectangle shape that exactly matches the chart's plot area using AddShapeInChartByScale (0‑100% coordinates). The shape is styled with a light‑red fill and red border before the workbook is saved as ChartWithExternalShape.xlsx.
    class Program
    {
        static void Main()
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
                chart.SetChartDataRange("A1:B4", true);
                chart.Calculate(); // Ensure the chart is calculated

                // Add a rectangle shape that exactly covers the plot area using scale coordinates (0%‑100%)
                Shape rectangle = chart.Shapes.AddShapeInChartByScale(
                    MsoDrawingType.Rectangle, // Shape type
                    PlacementType.Move,       // Placement behavior
                    0.0,                      // left (0% of chart width)
                    0.0,                      // top (0% of chart height)
                    1.0,                      // right (100% of chart width)
                    1.0);                     // bottom (100% of chart height)

                // Style the shape
                rectangle.Fill.SolidFill.Color = Color.FromArgb(255, 255, 200, 200); // Light red fill
                rectangle.Line.SolidFill.Color = Color.Red;                         // Red border

                // Save the workbook
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ChartWithExternalShape.xlsx");
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
