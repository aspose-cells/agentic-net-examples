// Title: Export an Excel chart to a high‑resolution JPEG file using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export chart as 300 DPI JPEG for PowerPoint | how to increase chart image resolution with Aspose.Cells ImageOrPrintOptions | save Excel chart to high resolution JPEG using chart.ToImage method | Aspose.Cells example for setting DPI when exporting chart to image | export first worksheet chart to JPEG with custom DPI in .NET
// Tags: Aspose.Cells chart JPEG export | ImageOrPrintOptions high DPI | C# chart rendering to image | Excel chart export for PowerPoint | Aspose.Cells ToImage method

using System;
using System.Drawing.Imaging;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The sample loads a workbook, retrieves the first chart, configures ImageOrPrintOptions with 300 DPI horizontal and vertical resolution, and exports the chart directly to a JPEG file suitable for high‑quality PowerPoint slides.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "ChartHighRes.jpg";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook that contains the chart
            Workbook workbook = new Workbook(inputPath);

            // Ensure the workbook has at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("The workbook does not contain any worksheets.");
                return;
            }

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet has at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the first worksheet.");
                return;
            }

            // Retrieve the first chart on the worksheet (adjust index if needed)
            Chart chart = worksheet.Charts[0];

            // Configure image options for high‑resolution output
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                HorizontalResolution = 300, // DPI – increase for higher resolution
                VerticalResolution = 300,   // DPI – increase for higher resolution
                Transparent = false         // Opaque background
                // ImageFormat is inferred from the output file extension
            };

            // Export the chart directly to a file using the configured options
            chart.ToImage(outputPath, imgOptions);

            Console.WriteLine($"Chart exported successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
