// Title: Export a specific Excel chart to a 300 DPI PNG image using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an Excel workbook, retrieves a chart by its index, configures ImageOrPrintOptions for 300 DPI, and saves the chart as a PNG file with Aspose.Cells. | Show how to set ImageOrPrintOptions.DpiX and DpiY to 300 and apply those options when calling Chart.ToImage to produce a high‑resolution PNG in a .NET application.
// Common Searches: Aspose.Cells C# export chart to PNG with 300 DPI resolution | How to increase image DPI when saving Excel chart using Aspose.Cells .NET | C# code example for high‑resolution chart export from workbook | Set custom DPI for chart image output in Aspose.Cells | Export specific chart from Excel to high‑resolution PNG in .NET
// Tags: chart export PNG custom DPI Aspose.Cells | ImageOrPrintOptions DpiX DpiY .NET | high resolution chart image Aspose.Cells | save Excel chart as PNG with 300 DPI C# | configure chart image resolution Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ExportChartApp
{
    // The example loads 'input.xlsx', verifies its presence, accesses the first worksheet, retrieves the first chart, and writes it to 'chart.png' using Chart.ToImage. By default the image uses the library’s standard DPI; to generate a 300 DPI PNG you would create an ImageOrPrintOptions object, set DpiX and DpiY to 300, and pass it to Chart.ToImage.
    class ExportChart
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

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one chart
                if (sheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found in the worksheet.");
                    return;
                }

                // Get the first chart on the sheet
                Chart chart = sheet.Charts[0];

                // Export the chart directly to a PNG file (default format is PNG)
                chart.ToImage(outputPath);

                Console.WriteLine($"Chart exported successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
