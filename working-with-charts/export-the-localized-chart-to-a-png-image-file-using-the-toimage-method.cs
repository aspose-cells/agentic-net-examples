// Title: Export the first chart from an Excel worksheet to a high‑resolution PNG using Aspose.Cells Chart.ToImage in C#
// Common Searches: c# how to save an Excel chart as a PNG with Aspose.Cells | using Aspose.Cells Chart.ToImage to generate high‑resolution PNG images | export chart from first worksheet to PNG file with 300 dpi in C# | Aspose.Cells ImageOrPrintOptions settings for PNG chart export
// Tags: Aspose.Cells Chart.ToImage PNG export | C# high‑resolution chart image generation | ImageOrPrintOptions 300dpi PNG configuration | first worksheet chart extraction Aspose.Cells | save Excel chart as PNG file

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace ExportChartToPng
{
    // The sample loads an Excel workbook, retrieves the first chart on the first worksheet, configures ImageOrPrintOptions for 300 dpi PNG output, ensures the target directory exists, and writes the chart to a PNG file using the Chart.ToImage method.
    class Program
    {
        static void Main(string[] args)
        {
            // Determine the input workbook path
            string inputPath = args.Length > 0 ? args[0] : "input.xlsx";
            string outputPath = "chart_output.png";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" does not exist.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Assume the chart we want to export is the first chart in the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                if (sheet.Charts.Count == 0)
                {
                    Console.WriteLine("Error: No charts found in the first worksheet.");
                    return;
                }

                Chart chart = sheet.Charts[0];

                // Configure image options for PNG output
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Png,
                    HorizontalResolution = 300,
                    VerticalResolution = 300,
                    OnePagePerSheet = true
                };

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Export the chart to a PNG file
                chart.ToImage(outputPath, imgOptions);

                Console.WriteLine($"Chart successfully exported to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An exception occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
