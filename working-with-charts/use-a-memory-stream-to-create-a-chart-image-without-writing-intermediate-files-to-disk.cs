// Title: Generate an Excel column chart and render it to a PNG MemoryStream using Aspose.Cells for .NET
// AI Prompts: Create a column chart from worksheet data and write the PNG image directly to a MemoryStream with Aspose.Cells. | Change the ImageOrPrintOptions to produce a JPEG stream instead of PNG when rendering the chart. | Add another data series to the chart, render the combined chart to a stream, and send the image bytes in an ASP.NET response.
// Common Searches: Aspose.Cells C# generate chart image in memory without file | How to export Aspose.Cells chart as JPEG byte array in C# | Return Excel chart image bytes from Aspose.Cells in a Web API | Create column chart PNG bytes from workbook using Aspose.Cells | Render Aspose.Cells chart to stream for ASP.NET response
// Tags: chart rendering to MemoryStream Aspose.Cells | column chart PNG output .NET | in‑memory chart image generation Aspose.Cells | export chart bytes C# Aspose.Cells | chart rendering without file system Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// Shows how to build a workbook, add a column chart, and use Aspose.Cells to render the chart directly into a PNG MemoryStream, avoiding any temporary files.
class ChartToMemoryStreamExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook entirely in memory
            var workbook = new Workbook();

            // Get the first worksheet
            var sheet = workbook.Worksheets[0];

            // Fill the worksheet with sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
            var chart = sheet.Charts[chartIdx];
            chart.Title.Text = "Sample Column Chart";

            // Define the series (values). Category data is optional for this simple example.
            int seriesIdx = chart.NSeries.Add("B2:B4", true);
            // If needed, you can set category data like this:
            // chart.NSeries[seriesIdx].CategoryData = "A2:A4";

            // Render the chart to a memory stream (PNG format)
            using (MemoryStream ms = new MemoryStream())
            {
                var imgOptions = new ImageOrPrintOptions();
                // The default image format is PNG, so we can omit setting ImageFormat explicitly.
                // imgOptions.ImageFormat = ImageFormat.Png; // Not required

                try
                {
                    chart.ToImage(ms, imgOptions);
                }
                catch (Exception renderEx)
                {
                    Console.WriteLine($"Error rendering chart: {renderEx.Message}");
                    return;
                }

                // The memory stream now contains the chart image bytes
                byte[] imageBytes = ms.ToArray();

                // Example usage: output the size of the generated image
                Console.WriteLine($"Chart image generated, size = {imageBytes.Length} bytes");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
