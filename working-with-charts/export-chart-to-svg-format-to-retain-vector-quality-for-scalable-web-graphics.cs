// Title: Export an Excel chart to SVG with Aspose.Cells for .NET (C#) while preserving vector quality
// Common Searches: C# Aspose.Cells export chart to SVG for web display | How to save an Excel chart as scalable SVG using Aspose.Cells .NET | Aspose.Cells chart.ToImage SVG output example in C# | Convert first worksheet chart to vector SVG with Aspose.Cells | Export Excel chart to SVG file using ImageOrPrintOptions in C#
// Tags: Aspose.Cells chart SVG export | C# export Excel chart to vector | chart.ToImage method SVG | ImageOrPrintOptions SVG configuration | save Excel chart as scalable SVG

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The sample loads an Excel workbook, checks for a chart on the first worksheet, and uses Aspose.Cells' ImageOrPrintOptions with the chart.ToImage method to write the chart directly to an SVG file, ensuring vector quality and handling missing files or runtime errors.
class ExportChartToSvg
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "chart.svg";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook containing the chart
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart
            Chart chart = worksheet.Charts[0];

            // Configure image options for SVG export (format inferred from file extension)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();

            // Export the chart directly to the SVG file
            chart.ToImage(outputPath, imgOptions);

            Console.WriteLine($"Chart successfully exported to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
