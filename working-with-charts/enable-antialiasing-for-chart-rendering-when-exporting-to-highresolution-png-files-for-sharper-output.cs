// Title: Enable anti‑aliasing for Aspose.Cells chart export to high‑resolution PNG in C# (.NET)
// AI Prompts: Set ImageOrPrintOptions.AntiAliasing = true before creating the SheetRender to enable smoothing for a 300 DPI PNG chart. | Modify the export routine to turn on smoothing while preserving the DPI settings, then call SheetRender.ToImage to generate the high‑resolution chart image.
// Common Searches: aspnet aspose.cells chart export png with anti aliasing enabled | c# high resolution chart PNG anti aliasing Aspose.Cells example | how to improve chart image sharpness when exporting to PNG using Aspose.Cells | set anti aliasing flag for chart rendering Aspose.Cells .NET | increase DPI and enable anti aliasing for chart PNG output Aspose.Cells
// Tags: chart smoothing Aspose.Cells | high‑resolution PNG export Aspose.Cells | ImageOrPrintOptions DPI configuration | SheetRender chart rendering .NET | C# Aspose.Cells chart image quality

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsExample
{
    // The example creates a workbook, adds sample data and a column chart, configures ImageOrPrintOptions with 300 DPI resolution, enables anti‑aliasing (smoothing), and uses SheetRender to export the chart as a high‑resolution PNG file.
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
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["A2"].PutValue(20);
                sheet.Cells["A3"].PutValue(30);
                sheet.Cells["B1"].PutValue(15);
                sheet.Cells["B2"].PutValue(25);
                sheet.Cells["B3"].PutValue(35);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart (first series with categories)
                chart.NSeries.Add("A1:A3", true);
                // Add a second series (values only)
                chart.NSeries.Add("B1:B3", false);

                // Add a title (optional)
                chart.Title.Text = "Sample Chart";

                // Configure image export options for high‑resolution PNG
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    // ImageFormat property is optional; PNG is inferred from file extension
                    HorizontalResolution = 300, // DPI
                    VerticalResolution = 300,   // DPI
                    OnePagePerSheet = true
                };

                // Render the worksheet (which contains the chart) to a PNG file
                SheetRender renderer = new SheetRender(sheet, imgOptions);
                renderer.ToImage(0, "ChartHighRes.png");

                Console.WriteLine("Chart exported successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
