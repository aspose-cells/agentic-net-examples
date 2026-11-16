// Title: Export an Excel chart to SVG with selectable text using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export chart to SVG preserving selectable text | How to save an Excel chart as scalable SVG with Aspose.Cells | Aspose.Cells Chart.ToImage SVG output with transparent background in .NET | Generate SVG from workbook chart using ImageOrPrintOptions SaveFormat.Svg | Export first worksheet chart to SVG file using Aspose.Cells C# example
// Tags: chart to svg Aspose.Cells C# | ImageOrPrintOptions SaveFormat.Svg usage | transparent background svg chart export | one page per sheet svg rendering | selectable text in exported SVG chart | Aspose.Cells chart rendering options

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The sample loads an Excel workbook, checks for a chart on the first worksheet, configures ImageOrPrintOptions with SaveFormat.Svg, OnePagePerSheet, and Transparent settings, and then uses Chart.ToImage to export the chart as an SVG file that retains selectable text for web display.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "chart.svg";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook containing the chart
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("Error: No charts found in the first worksheet.");
                return;
            }

            // Retrieve the first chart on the worksheet
            Chart chart = worksheet.Charts[0];

            // Configure image options for SVG output
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                SaveFormat = SaveFormat.Svg,   // Export as SVG for scalable, selectable text
                OnePagePerSheet = true,
                Transparent = true
            };

            // Export the chart to an SVG file
            chart.ToImage(outputPath, imgOptions);
            Console.WriteLine($"Chart exported successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
