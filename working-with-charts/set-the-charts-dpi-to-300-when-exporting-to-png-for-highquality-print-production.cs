// Title: Export an Excel chart to a 300 DPI PNG file using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a workbook, retrieves the first chart, configures ImageOrPrintOptions with 300 dpi for both horizontal and vertical resolution, and saves the chart as a PNG image. | Show how to implement error handling for a missing input workbook and for worksheets that contain no charts before exporting a high‑resolution chart image. | Demonstrate using Chart.ToImage together with ImageOrPrintOptions to produce a high‑resolution PNG export of an Excel chart.
// Common Searches: Aspose.Cells export chart to PNG with 300 DPI in C# | C# set chart image resolution when saving as PNG using Aspose.Cells | How to increase DPI of chart image exported from Excel with Aspose.Cells | ImageOrPrintOptions HorizontalResolution 300 for chart export Aspose.Cells .NET
// Tags: Aspose.Cells chart export PNG high DPI | ImageOrPrintOptions resolution settings C# | Chart.ToImage high resolution export | Excel chart PNG 300 DPI Aspose.Cells | C# error handling missing workbook chart

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The example loads 'input.xlsx', checks for the presence of a chart, sets ImageOrPrintOptions HorizontalResolution and VerticalResolution to 300 dpi, and uses Chart.ToImage to export the first chart as a high‑quality PNG file named 'chart.png', with error handling for missing files and absent charts.
class ExportChartWithHighDpi
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

            // Load the workbook containing the chart
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
                HorizontalResolution = 300,
                VerticalResolution = 300
                // Image format is inferred from the output file extension (.png)
            };

            // Export the chart directly to a PNG file using the specified DPI
            chart.ToImage(outputPath, imgOptions);

            Console.WriteLine($"Chart exported to PNG with 300 DPI at '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
