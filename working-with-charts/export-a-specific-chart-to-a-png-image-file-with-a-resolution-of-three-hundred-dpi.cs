// Title: Export a chart from an Excel workbook to a 300 DPI PNG file using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook, selects a chart, and saves it as a PNG image with 300 DPI using Aspose.Cells. | Show how to configure ImageOrPrintOptions.HorizontalResolution and VerticalResolution to 300 and export the chart with a transparent background.
// Common Searches: Aspose.Cells C# export Excel chart as 300 DPI PNG image | How to set DPI when saving a chart to PNG with Aspose.Cells | C# code to export first worksheet chart to high‑resolution PNG using Aspose.Cells | Saving Excel chart with transparent background as PNG in .NET | ImageOrPrintOptions DPI configuration for chart export Aspose.Cells
// Tags: export chart to PNG with custom DPI Aspose.Cells | ImageOrPrintOptions set 300 DPI | C# Aspose.Cells chart image export | transparent background PNG chart Aspose.Cells | high‑resolution chart image generation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The program loads 'input.xlsx', retrieves the first chart from the first worksheet, sets ImageOrPrintOptions to 300 DPI for both dimensions (optionally enabling a transparent background), and saves the chart as 'chart.png' using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "chart.png";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook containing the chart
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index or name as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Retrieve the first chart on the sheet
            Chart chart = worksheet.Charts[0];

            // Configure image export options with 300 DPI resolution
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                HorizontalResolution = 300,
                VerticalResolution = 300,
                Transparent = true // optional: make background transparent
            };

            // Export the chart to a PNG file using the defined options
            chart.ToImage(outputPath, imgOptions);

            Console.WriteLine($"Chart exported successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
