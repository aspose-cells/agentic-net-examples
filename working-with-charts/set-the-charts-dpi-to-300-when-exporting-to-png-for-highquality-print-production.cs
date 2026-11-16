// Title: Export an Excel chart to a 300 DPI PNG file using Aspose.Cells for .NET (C#)
// Common Searches: how to export an Excel chart as a 300 DPI PNG with Aspose.Cells in C# | Aspose.Cells C# set chart image resolution to 300 DPI when saving to PNG | C# ImageOrPrintOptions HorizontalResolution VerticalResolution example for chart export | export high‑resolution chart image from workbook using Aspose.Cells ToImage method | generate print‑ready PNG of Excel chart at 300 DPI with Aspose.Cells
// Tags: chart export high DPI PNG Aspose.Cells | ImageOrPrintOptions set resolution C# | Aspose.Cells ToImage method high quality output | export Excel chart as 300 DPI image | C# generate print‑ready chart PNG | Aspose.Cells chart rendering with custom DPI

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The sample loads an Excel workbook, retrieves the first chart, configures ImageOrPrintOptions with 300 DPI horizontal and vertical resolution, and uses Chart.ToImage to export the chart as a high‑quality PNG suitable for print production.
class ExportChartWithHighDpi
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "ChartExport.png";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook that contains the chart
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("Error: No charts found in the first worksheet.");
                return;
            }

            // Get the first chart on the worksheet
            Chart chart = worksheet.Charts[0];

            // Configure image export options with 300 DPI for high‑quality PNG
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                HorizontalResolution = 300, // DPI horizontally
                VerticalResolution = 300    // DPI vertically
            };

            // Export the chart to a PNG file using the specified DPI
            chart.ToImage(outputPath, imgOptions);

            Console.WriteLine($"Chart exported successfully with 300 DPI to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
