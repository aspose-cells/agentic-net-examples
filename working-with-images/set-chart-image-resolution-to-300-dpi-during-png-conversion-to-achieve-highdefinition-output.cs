// Title: Export an Aspose.Cells chart to a 300 DPI PNG image with C#
// AI Prompts: Write C# code that creates a chart in an Aspose.Cells workbook and saves it as a PNG with 300 DPI using ImageOrPrintOptions. | Show how to set HorizontalResolution and VerticalResolution to 300 DPI when exporting a specific chart to PNG in Aspose.Cells. | Provide a snippet that configures ImageOrPrintOptions for high‑definition PNG output and calls chart.ToImage in .NET.
// Common Searches: Aspose.Cells C# export chart as PNG with custom DPI | How to set 300 DPI resolution for chart images in Aspose.Cells | ImageOrPrintOptions 300 DPI PNG example for .NET | Saving Excel chart to high‑resolution PNG using Aspose.Cells | C# code to generate high‑definition chart PNG from workbook
// Tags: chart PNG high DPI export | configure DPI for chart image C# | high‑definition chart export Aspose.Cells | export specific chart to PNG .NET | set chart resolution Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The example creates a workbook, adds sample data and a column chart, configures ImageOrPrintOptions with 300 DPI horizontal and vertical resolution, and exports the chart directly to a PNG file named Chart300DPI.png.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and populate sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);
            sheet.Cells["B1"].PutValue(15);
            sheet.Cells["B2"].PutValue(25);
            sheet.Cells["B3"].PutValue(35);

            // Add a column chart to the worksheet
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIdx];
            chart.NSeries.Add("A1:A3", true);
            // Category data line removed because the property is not available in this version

            // Configure image options for 300 DPI PNG output
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                HorizontalResolution = 300,
                VerticalResolution = 300,
                OnePagePerSheet = true
            };

            // Export the chart directly to a high‑definition PNG file
            string outputPath = "Chart300DPI.png";
            chart.ToImage(outputPath, imgOptions);

            Console.WriteLine($"Chart exported successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
