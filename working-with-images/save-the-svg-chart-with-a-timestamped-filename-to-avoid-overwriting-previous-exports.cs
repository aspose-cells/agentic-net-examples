// Title: Save an Excel chart as a timestamped SVG file with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export chart to SVG with unique timestamped filename | How to prevent overwriting SVG chart exports when using Aspose.Cells in .NET | Generate SVG image from an Excel chart using Aspose.Cells and include date‑time in the file name | Aspose.Cells chart.ToImage save as SVG using DateTime.Now naming pattern | Export first worksheet chart to SVG file with timestamp using Aspose.Cells API
// Tags: Aspose.Cells chart SVG export options | C# create unique file names for exported assets | ImageOrPrintOptions SaveFormat.Svg configuration | prevent overwriting exported SVG files Aspose.Cells | export Excel chart as SVG using .NET API

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The example loads a workbook, checks for a chart on the first worksheet, builds a filename that includes the current date and time (e.g., Chart_20231127_154530.svg), configures ImageOrPrintOptions with SaveFormat.Svg, and uses chart.ToImage to export the chart to a timestamped SVG file, ensuring previous exports are not overwritten and handling missing files or other exceptions.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";

            // Ensure the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Verify that the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart in the worksheet
            Chart chart = sheet.Charts[0];

            // Build a timestamped filename to avoid overwriting previous exports
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string fileName = $"Chart_{timestamp}.svg";

            // Configure SVG export options
            ImageOrPrintOptions options = new ImageOrPrintOptions();
            options.SaveFormat = SaveFormat.Svg;
            options.OnePagePerSheet = true; // optional setting

            // Export the chart to an SVG file
            chart.ToImage(fileName, options);

            Console.WriteLine($"Chart successfully exported to: {fileName}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
