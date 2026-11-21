// Title: Export an Excel chart to a high‑resolution 300 DPI PNG file using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export chart to PNG with 300 DPI resolution | How to set horizontal and vertical DPI for chart image in Aspose.Cells | Saving an Excel chart as a high‑resolution PNG using ImageOrPrintOptions in .NET | C# code to render the first worksheet chart to a 300 DPI PNG with Aspose.Cells | Export chart from workbook to PNG at 300 DPI using Aspose.Cells API
// Tags: chart.ToImage ImageOrPrintOptions DPI | ImageOrPrintOptions horizontalResolution Aspose.Cells | Aspose.Cells chart image rendering PNG | high‑resolution chart export .NET | Excel workbook chart extraction PNG

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The example loads a workbook, checks for a chart on the first worksheet, configures ImageOrPrintOptions with 300 DPI horizontal and vertical resolution, and uses Chart.ToImage to export the chart as a PNG file.
class ExportChartToPng
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
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook containing the chart
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet has at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart on the worksheet
            Chart chart = worksheet.Charts[0];

            // Configure image export options (300 DPI PNG)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // Default format is PNG; set resolution via horizontal/vertical DPI
                HorizontalResolution = 300,
                VerticalResolution = 300
            };

            // Export the chart to a PNG file
            chart.ToImage(outputPath, imgOptions);
            Console.WriteLine($"Chart exported successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
