// Title: Convert an Excel workbook to a 24‑bit multi‑page TIFF file with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export Excel worksheet to 24-bit TIFF image | How to generate high‑resolution multi‑page TIFF from a workbook using Aspose.Cells | Set TIFF color depth to 24‑bit with ImageOrPrintOptions in Aspose.Cells | Apply LZW compression and 300 DPI when converting Excel to TIFF in .NET | Render entire workbook to a single TIFF file using WorkbookRender Aspose.Cells
// Tags: Aspose.Cells ImageOrPrintOptions TIFF conversion | 24-bit TIFF color depth Aspose.Cells | WorkbookRender multi-page TIFF output | LZW compression TIFF Aspose.Cells | 300 DPI TIFF resolution Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace ConvertWorkbookToTiff
{
    // The example loads an Excel workbook, configures ImageOrPrintOptions to use TIFF format with LZW compression, 24‑bit color depth, and 300 DPI resolution, then uses WorkbookRender to render the whole workbook into a single multi‑page TIFF file, handling missing files and runtime errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths.
            // You can modify these paths as needed.
            string inputPath = "input.xlsx";
            string outputPath = "output_24bit.tiff";

            // Verify that the input file exists.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The input file \"{inputPath}\" does not exist.");
                return;
            }

            try
            {
                // Load the workbook from the specified file.
                Workbook workbook = new Workbook(inputPath);

                // Configure image rendering options.
                ImageOrPrintOptions options = new ImageOrPrintOptions
                {
                    // Set the output format to TIFF.
                    ImageType = ImageType.Tiff,

                    // Use a lossless compression (LZW) for better quality.
                    TiffCompression = TiffCompression.CompressionLZW,

                    // Set the desired color depth to 24‑bit.
                    TiffColorDepth = ColorDepth.Format24bpp,

                    // Optional: set resolution for higher quality output.
                    HorizontalResolution = 300,
                    VerticalResolution = 300
                };

                // Create a WorkbookRender object with the workbook and options.
                WorkbookRender renderer = new WorkbookRender(workbook, options);

                // Render the entire workbook to a single multi‑page TIFF file.
                renderer.ToImage(outputPath);

                Console.WriteLine($"Successfully converted \"{inputPath}\" to 24‑bit TIFF \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Output any errors that occur during processing.
                Console.WriteLine("An error occurred while converting the workbook to TIFF:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
