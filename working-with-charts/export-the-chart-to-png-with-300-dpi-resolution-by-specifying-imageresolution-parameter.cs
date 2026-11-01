// Title: Export an Excel chart to a high‑resolution 300 DPI PNG using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export chart as PNG with 300 DPI | How to set image resolution for chart export in Aspose.Cells .NET | C# ImageOrPrintOptions 300 DPI chart image from workbook | Saving Excel chart to high‑resolution PNG using Aspose.Cells | chart.ToImage method with custom DPI in Aspose.Cells
// Tags: chart export to PNG with custom DPI | ImageOrPrintOptions resolution settings | Aspose.Cells high‑resolution chart image | C# chart.ToImage PNG output | Excel workbook chart extraction Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The sample loads an Excel workbook, retrieves the first chart, configures ImageOrPrintOptions with 300 DPI horizontal and vertical resolution, ensures the output folder exists, and exports the chart to a PNG file using chart.ToImage.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "chart.png";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook containing the chart
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart on the worksheet (adjust index if needed)
            Chart chart = sheet.Charts[0];

            // Configure image export options with 300 DPI resolution
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // Image format is inferred from the file extension, so no need to set ImageFormat explicitly
                HorizontalResolution = 300, // 300 DPI horizontally
                VerticalResolution = 300    // 300 DPI vertically
            };

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Export the chart to a PNG file using the specified resolution
            chart.ToImage(outputPath, imgOptions);
            Console.WriteLine($"Chart exported successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
