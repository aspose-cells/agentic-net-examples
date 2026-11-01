// Title: Export an Excel chart to a high‑resolution JPEG file for PowerPoint using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells export chart to JPEG with 300 DPI resolution | How to save an Excel chart as a high quality JPEG for PowerPoint slides using Aspose.Cells | Aspose.Cells ImageOrPrintOptions set DPI when converting chart to image | Export first worksheet chart to JPEG file in .NET | Generate high‑resolution chart image from Excel workbook using Aspose.Cells API
// Tags: Aspose.Cells chart JPEG export | high‑resolution chart image .NET | ImageOrPrintOptions DPI setting | Excel chart to image for PowerPoint | C# Aspose.Cells chart rendering

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using System.Drawing;

// Loads an Excel workbook, retrieves the first chart from the first worksheet, configures ImageOrPrintOptions with 300 DPI horizontal and vertical resolution, and saves the chart as a high‑resolution JPEG file suitable for embedding in PowerPoint presentations.
class ExportChartToJpeg
{
    static void Main()
    {
        try
        {
            // Path to the source Excel file containing the chart
            string workbookPath = @"C:\Data\SampleChart.xlsx";

            // Path where the high‑resolution JPEG image will be saved
            string jpegPath = @"C:\Data\ChartImage.jpg";

            // Verify that the workbook file exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook file not found: {workbookPath}");
                return;
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(jpegPath);
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the first worksheet.");
                return;
            }

            // Retrieve the first chart
            Chart chart = sheet.Charts[0];

            // Configure image export options for high resolution
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // Image format is inferred from the file extension (.jpg)
                HorizontalResolution = 300, // DPI
                VerticalResolution = 300,   // DPI
                Transparent = false
            };

            // Export the chart to a high‑resolution JPEG image
            chart.ToImage(jpegPath, imgOptions);

            Console.WriteLine($"Chart exported successfully to: {jpegPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
