// Title: Export an Excel chart to a 300 DPI PNG image using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export chart as PNG with 300 DPI resolution | set horizontal and vertical resolution for chart image in Aspose.Cells | how to generate high‑definition PNG from Excel chart using ImageOrPrintOptions | C# code to save Excel chart to high‑resolution PNG file | increase chart image DPI when converting Excel to PNG with Aspose.Cells
// Tags: chart export PNG Aspose.Cells | ImageOrPrintOptions DPI setting | high‑resolution chart image C# | Excel chart to PNG with custom resolution | Aspose.Cells chart image quality

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// Loads an Excel workbook, retrieves the first chart, and saves it as a PNG file with 300 DPI horizontal and vertical resolution using Aspose.Cells ImageOrPrintOptions.
class ChartExport
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
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("Error: No charts found in the first worksheet.");
                return;
            }

            // Get the first chart
            Chart chart = sheet.Charts[0];

            // Configure image options for high‑definition PNG (300 DPI)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                HorizontalResolution = 300,
                VerticalResolution = 300
                // Image format is inferred from the output file extension (.png)
            };

            // Export the chart to a PNG file with the specified DPI
            chart.ToImage(outputPath, imgOptions);

            Console.WriteLine($"Chart exported to '{outputPath}' at 300 DPI.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
