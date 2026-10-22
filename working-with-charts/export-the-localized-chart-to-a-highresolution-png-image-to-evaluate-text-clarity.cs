// Title: Export a localized Excel chart to a 300 DPI PNG with transparent background using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an existing .xlsx file, verifies the first worksheet contains a chart, and saves that chart as a 300 DPI PNG with a transparent background using Aspose.Cells. | Show how to configure ImageOrPrintOptions.HorizontalResolution and VerticalResolution to 300 DPI and enable transparency before calling Chart.ToImage in Aspose.Cells. | Write a robust C# snippet that checks for file existence and chart count, then exports the first chart to a high‑resolution PNG while handling possible exceptions.
// Common Searches: Aspose.Cells C# export chart to PNG with 300 DPI resolution | How to save Excel chart as high‑resolution PNG using Aspose.Cells .NET | Set transparent background when converting Excel chart to image with Aspose.Cells | Check if worksheet contains charts before exporting image in Aspose.Cells
// Tags: chart to PNG high‑resolution Aspose.Cells | ImageOrPrintOptions DPI setting .NET | export Excel chart transparent background | validate chart existence Aspose.Cells | load workbook and render chart image

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// Loads an .xlsx workbook, verifies the first worksheet has a chart, configures ImageOrPrintOptions for 300 DPI and transparency, and exports the chart to a PNG file using Aspose.Cells.
class ExportChart
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "chart_highres.png";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook that contains the chart
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("Error: No charts found in the worksheet.");
                return;
            }

            // Retrieve the first chart on the sheet
            Chart chart = worksheet.Charts[0];

            // Configure image options for high‑resolution PNG (e.g., 300 DPI)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                HorizontalResolution = 300, // DPI
                VerticalResolution = 300,   // DPI
                Transparent = true          // optional: keep background transparent
            };

            // Export the chart to a PNG file with the specified resolution
            chart.ToImage(outputPath, imgOptions);
            Console.WriteLine($"Chart exported successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
