// Title: Export a chart from Aspose.Cells to a 300 DPI PNG image using ImageOrPrintOptions in C#
// AI Prompts: Write C# code that creates a workbook, adds a column chart, and saves the chart as a PNG file with 300 DPI using ImageOrPrintOptions. | Show how to adjust HorizontalResolution and VerticalResolution in ImageOrPrintOptions to control the DPI of a chart image exported from Aspose.Cells. | Provide a snippet that exports an Aspose.Cells chart to PNG while allowing the DPI value to be changed dynamically.
// Common Searches: Aspose.Cells export chart to PNG with specific DPI in C# | How to set image resolution for chart export using ImageOrPrintOptions in Aspose.Cells | C# code sample for saving Aspose.Cells chart as high‑resolution PNG | Changing horizontal and vertical resolution when converting Aspose.Cells chart to image
// Tags: Aspose.Cells chart export PNG | ImageOrPrintOptions DPI setting | C# chart to high‑resolution image | Aspose.Cells ToImage method resolution | export chart with 300 DPI Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The program creates a workbook with sample data, adds a column chart, and uses ImageOrPrintOptions to export the chart as a PNG file named 'Chart300DPI.png' with both horizontal and vertical resolutions set to 300 DPI.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
            var chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";      // Categories

            // Define image export options with 300 DPI resolution
            var imgOptions = new ImageOrPrintOptions
            {
                // Default format is PNG; explicit format setting omitted to avoid API mismatch
                HorizontalResolution = 300, // DPI
                VerticalResolution = 300    // DPI
            };

            // Export the chart to a PNG file using the specified resolution
            using (var stream = new MemoryStream())
            {
                chart.ToImage(stream, imgOptions);
                File.WriteAllBytes("Chart300DPI.png", stream.ToArray());
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
