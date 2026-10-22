// Title: Create a column chart with Aspose.Cells and render it to a MemoryStream as PNG in C#
// AI Prompts: Generate a column chart from worksheet data and write the chart image to a MemoryStream using Aspose.Cells ImageOrPrintOptions in C#. | Change the rendering options to produce a JPEG stream and return the resulting byte array from the chart. | Add proper disposal and comprehensive error handling around the MemoryStream when converting a chart to an image.
// Common Searches: aspocells render chart to memory stream c# example | how to get chart image bytes from Aspose.Cells without saving to disk | c# Aspose.Cells column chart to png in memory | convert Aspose.Cells chart to byte array using MemoryStream
// Tags: Aspose.Cells chart rendering to MemoryStream | C# column chart PNG generation with ImageOrPrintOptions | in‑memory chart image conversion Aspose.Cells | avoid disk I/O when exporting chart Aspose.Cells | retrieve chart byte array from workbook C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;   // Required for ImageOrPrintOptions

// The example builds a workbook in memory, fills it with sample data, adds a column chart, and uses ImageOrPrintOptions to render the chart directly into a MemoryStream as a PNG, then extracts the image bytes for further use.
class ChartToMemoryStreamExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook in memory
            using (Workbook workbook = new Workbook())
            {
                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart (rows 5‑20, columns 0‑10)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set data range for series and categories
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Prepare image options (default format is PNG)
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();

                // Render chart to a memory stream
                using (MemoryStream chartStream = new MemoryStream())
                {
                    chart.ToImage(chartStream, imgOptions);
                    chartStream.Position = 0; // Reset for reading

                    // Obtain byte array of the PNG image
                    byte[] imageBytes = chartStream.ToArray();
                    Console.WriteLine($"Chart image generated, size = {imageBytes.Length} bytes.");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
