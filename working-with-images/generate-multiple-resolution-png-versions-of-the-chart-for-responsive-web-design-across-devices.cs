// Title: Export an Excel chart to PNG files at 72, 150, and 300 DPI using Aspose.Cells for .NET
// Common Searches: how to render an Excel chart to PNG with specific DPI using Aspose.Cells C# | Aspose.Cells generate chart images for different screen resolutions | C# code to export a workbook chart to multiple PNG files with varying DPI | set horizontal and vertical resolution when saving a chart as PNG with Aspose.Cells
// Tags: chart to PNG Aspose.Cells | set image resolution Aspose.Cells | export chart multiple DPI .NET | ImageOrPrintOptions PNG output | responsive chart image generation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace ChartResponsiveExport
{
    // The example loads an Excel workbook, checks for a chart on the first worksheet, and uses ImageOrPrintOptions with custom HorizontalResolution and VerticalResolution values (72, 150, 300 DPI) to render the chart to separate PNG files named chart_72dpi.png, chart_150dpi.png, and chart_300dpi.png. OnePagePerSheet ensures the whole chart fits on each image, and the workbook is disposed after processing.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel workbook that contains the chart.
            const string inputFile = "input.xlsx";

            // Verify that the input file exists.
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file \"{inputFile}\" was not found in the application directory.");
                return;
            }

            try
            {
                // Load the workbook.
                Workbook workbook = new Workbook(inputFile);

                // Access the first worksheet (you can change the index or name as needed).
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one chart.
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine("Error: No charts were found in the first worksheet.");
                    return;
                }

                // Retrieve the first chart. Adjust the index if you need a different chart.
                Chart chart = worksheet.Charts[0];

                // Define the DPI (resolution) values for which PNG images will be generated.
                int[] dpiValues = new int[] { 72, 150, 300 };

                foreach (int dpi in dpiValues)
                {
                    // Configure image rendering options.
                    ImageOrPrintOptions options = new ImageOrPrintOptions
                    {
                        ImageType = ImageType.Png,
                        HorizontalResolution = dpi,
                        VerticalResolution = dpi,
                        // Ensure the whole chart fits on a single image.
                        OnePagePerSheet = true
                    };

                    // Build the output file name (e.g., chart_72dpi.png).
                    string outputFileName = $"chart_{dpi}dpi.png";

                    // Render the chart to a PNG file using the specified options.
                    chart.ToImage(outputFileName, options);

                    Console.WriteLine($"Chart rendered at {dpi} DPI and saved to \"{outputFileName}\".");
                }

                // Optional: Dispose of the workbook if you are done with it.
                workbook.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
