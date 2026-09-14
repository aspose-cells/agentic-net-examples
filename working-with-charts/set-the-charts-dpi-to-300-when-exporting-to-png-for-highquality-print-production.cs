// Title: Export a chart to a 300 DPI PNG image with Aspose.Cells in C# for print‑quality output
// AI Prompts: Write C# code that creates a column chart using Aspose.Cells and saves it as a 300 DPI PNG by configuring ImageOrPrintOptions. | Show how to set ImageOrPrintOptions.HorizontalResolution and VerticalResolution to 300 for high‑resolution chart rendering in Aspose.Cells. | Provide a complete example that builds sample data, adds a chart, and exports it to a print‑ready PNG file at 300 DPI.
// Common Searches: how to export Aspose.Cells chart as 300 DPI PNG in C# | Aspose.Cells ImageOrPrintOptions set resolution for chart image | C# generate print‑ready chart PNG with 300 DPI using Aspose.Cells | set chart image DPI when saving to PNG with Aspose.Cells .NET | high resolution chart export Aspose.Cells example C#
// Tags: Aspose.Cells chart export PNG 300 DPI | ImageOrPrintOptions set resolution for chart | C# high‑resolution chart image generation | column chart PNG export Aspose.Cells | print‑ready chart image Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The program creates a workbook, adds sample data, inserts a column chart, configures ImageOrPrintOptions with 300 DPI horizontal and vertical resolution, and exports the chart to a PNG file named Chart300DPI.png.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and name it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Fill sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIdx];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure image export options with 300 DPI (default format is PNG)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                HorizontalResolution = 300, // DPI horizontally
                VerticalResolution = 300    // DPI vertically
            };

            // Export the chart to a PNG file using the high‑resolution settings
            using (MemoryStream ms = new MemoryStream())
            {
                chart.ToImage(ms, imgOptions);
                File.WriteAllBytes("Chart300DPI.png", ms.ToArray());
            }

            Console.WriteLine("Chart exported successfully to Chart300DPI.png");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
