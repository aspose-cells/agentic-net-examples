// Title: Export an Excel chart to a high‑resolution 300 DPI PNG file and save it next to the workbook using Aspose.Cells for .NET
// Common Searches: C# export Excel chart as 300 DPI PNG with Aspose.Cells | save chart image in same folder as workbook using Aspose.Cells .NET | set horizontal and vertical resolution for chart image export Aspose.Cells | how to use ImageOrPrintOptions to export chart to PNG in C# | Aspose.Cells export first worksheet chart to PNG file
// Tags: chart to PNG export Aspose.Cells | 300 DPI image export options C# | ImageOrPrintOptions resolution setting | chart image file placement next to workbook | Aspose.Cells chart image generation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The example loads an existing Excel workbook, retrieves the first chart from the first worksheet, configures ImageOrPrintOptions with 300 dpi horizontal and vertical resolution, and uses Chart.ToImage to export the chart as a PNG file named "Chart.png" in the same directory as the workbook. It then saves the workbook (if modified) and includes error handling for missing files and export failures.
class ExportChartExample
{
    static void Main()
    {
        // Path to the input workbook
        string workbookPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Input workbook not found: {workbookPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Access the first worksheet (adjust index as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one chart in the worksheet
            if (worksheet.Charts.Count > 0)
            {
                // Get the first chart
                Chart chart = worksheet.Charts[0];

                // Configure image export options for PNG at 300 DPI
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    HorizontalResolution = 300,
                    VerticalResolution = 300,
                    OnePagePerSheet = true
                };

                // Build the output file path (same folder as the workbook)
                string folder = Path.GetDirectoryName(Path.GetFullPath(workbookPath)) ?? Directory.GetCurrentDirectory();
                string chartImagePath = Path.Combine(folder, "Chart.png");

                try
                {
                    // Export the chart to PNG using the specified options
                    chart.ToImage(chartImagePath, imgOptions);
                    Console.WriteLine($"Chart exported to: {chartImagePath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to export chart: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("No charts found in the first worksheet.");
            }

            // Save the workbook (if any modifications were made)
            string outputPath = "output.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
