// Title: Export a chart from an Excel workbook to a 300 DPI PNG file using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export chart to PNG with 300 DPI resolution | How to set image resolution when saving an Excel chart as PNG in .NET | C# code to convert an Excel chart to a high‑resolution PNG using Aspose.Cells | Export first worksheet chart to PNG at 300 DPI with Aspose.Cells | ImageOrPrintOptions DPI setting for chart export in Aspose.Cells
// Tags: Aspose.Cells chart image DPI configuration | ImageOrPrintOptions set 300 DPI for PNG export | C# export Excel chart as high‑resolution PNG | chart ToImage method with resolution settings | export chart from workbook to PNG file

using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using System;
using System.IO;

// The example loads a workbook, checks for a chart on the first worksheet, configures ImageOrPrintOptions with 300 DPI horizontal and vertical resolution, and uses the chart's ToImage method to export the chart as a PNG image, handling missing files and export errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook containing the chart
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Retrieve the first chart
            Chart chart = worksheet.Charts[0];

            // Configure image export options with 300 DPI resolution
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                HorizontalResolution = 300,
                VerticalResolution = 300
                // Image format is inferred from the output file extension (PNG)
            };

            // Export the chart to a PNG file
            const string outputPath = "chart.png";

            try
            {
                chart.ToImage(outputPath, imgOptions);
                Console.WriteLine($"Chart exported successfully to '{outputPath}'.");
            }
            catch (Exception exportEx)
            {
                Console.WriteLine($"Failed to export chart: {exportEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
