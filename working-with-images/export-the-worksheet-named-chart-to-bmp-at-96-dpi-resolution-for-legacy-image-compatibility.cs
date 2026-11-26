// Title: Export the Chart worksheet to BMP files at 96 DPI using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells convert a specific worksheet named Chart to BMP with 96 DPI | How to use SheetRender to save each page of an Excel sheet as a BMP image in .NET | Set image resolution when exporting Excel worksheet to BMP with Aspose.Cells | Aspose.Cells render chart worksheet to multiple BMP files from command line | Export Excel worksheet to BMP format without metadata using Aspose.Cells C#
// Tags: SheetRender BMP export Aspose.Cells | ImageOrPrintOptions 96 DPI BMP | export chart worksheet as bitmap C# | multiple page BMP output from Excel sheet | Aspose.Cells legacy BMP image generation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace ExportWorksheetToBmp
{
    // The example loads a workbook, finds the worksheet named "Chart", configures ImageOrPrintOptions for BMP format with 96 DPI resolution, and uses SheetRender to generate a separate BMP file for each page, saving the images to a user‑specified directory.
    class Program
    {
        static void Main(string[] args)
        {
            // Validate arguments
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: ExportWorksheetToBmp <inputWorkbookPath> <outputDirectory>");
                return;
            }

            string inputPath = args[0];
            string outputDir = args[1];

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                return;
            }

            try
            {
                // Ensure output directory exists
                Directory.CreateDirectory(outputDir);

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Find the worksheet named "Chart"
                Worksheet chartSheet = null;
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    if (string.Equals(ws.Name, "Chart", StringComparison.OrdinalIgnoreCase))
                    {
                        chartSheet = ws;
                        break;
                    }
                }

                if (chartSheet == null)
                {
                    Console.WriteLine("Error: Worksheet named \"Chart\" was not found in the workbook.");
                    return;
                }

                // Configure image options for BMP at 96 DPI
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = Aspose.Cells.Drawing.ImageType.Bmp,
                    HorizontalResolution = 96,
                    VerticalResolution = 96,
                    // Legacy compatibility: do not embed any extra metadata
                    // (Bmp format is inherently simple, no extra settings required)
                };

                // Create a SheetRender for the target worksheet
                SheetRender sheetRender = new SheetRender(chartSheet, imgOptions);

                // Render each page of the worksheet to a separate BMP file
                for (int pageIndex = 0; pageIndex < sheetRender.PageCount; pageIndex++)
                {
                    string outputFile = Path.Combine(outputDir, $"Chart_page{pageIndex + 1}.bmp");
                    sheetRender.ToImage(pageIndex, outputFile);
                    Console.WriteLine($"Rendered page {pageIndex + 1} to \"{outputFile}\"");
                }

                // Clean up
                sheetRender.Dispose();

                Console.WriteLine("Export completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An exception occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
