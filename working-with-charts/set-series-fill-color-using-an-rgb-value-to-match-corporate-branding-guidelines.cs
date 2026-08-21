// Title: Set a chart series fill color with an RGB value in Aspose.Cells for .NET
// Description: Creates a workbook, adds a column chart with sample data, and applies a solid fill to the first series using the RGB values (79, 129, 189). The workbook is saved as SeriesFillColorDemo.xlsx, demonstrating how to match corporate branding colors programmatically.
// Keywords: Aspose.Cells chart series color | C# set series fill RGB | solid fill format Aspose.Cells | column chart series color .NET | Color.FromArgb Aspose.Cells example | branding colors Excel chart | programmatic chart styling
// Common Searches: how to change chart series color Aspose.Cells C# | set RGB fill for Excel chart series using Aspose.Cells | apply corporate brand colors to Aspose.Cells charts | solid fill format for chart series .NET | Aspose.Cells example for series fill color
// Developer Intent: Apply a specific RGB solid fill to the first series of a column chart generated with Aspose.Cells.
// Use Cases: Enforce corporate brand palettes in automatically generated Excel reports. | Create reusable chart templates with predefined series colors for consistency across dashboards. | Batch‑process workbooks to ensure all charts use the same visual style before distribution.
// AI Prompts: Generate code that assigns a different RGB color to each series in an Aspose.Cells chart. | Show how to apply a gradient fill to a chart series using Aspose.Cells for .NET. | Explain how to update the fill format of an existing chart series after loading a workbook with Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsSeriesFillColorDemo
{
    // Creates a workbook, adds a column chart with sample data, and applies a solid fill to the first series using the RGB values (79, 129, 189). The workbook is saved as SeriesFillColorDemo.xlsx, demonstrating how to match corporate branding colors programmatically.
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

                // Define the series data range
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Set the fill color of the first series (RGB: 79,129,189)
                Series series = chart.NSeries[0];
                series.Area.FillFormat.FillType = FillType.Solid; // Use solid fill
                series.Area.FillFormat.SolidFill.Color = Color.FromArgb(79, 129, 189);

                // Save the workbook
                workbook.Save("SeriesFillColorDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
