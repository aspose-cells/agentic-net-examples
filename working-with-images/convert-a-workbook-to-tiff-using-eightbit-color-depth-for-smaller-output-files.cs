// Title: Convert an Excel workbook to an 8‑bit multi‑page TIFF file using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# convert Excel to 8-bit TIFF with LZW compression | How to render a workbook as a multi‑page TIFF with 256 colors in .NET | Set TiffColorDepth to Format8bpp when exporting Excel to TIFF using Aspose | Create a single TIFF file from multiple worksheets with Aspose.Cells | C# code example for saving an Excel workbook as an 8‑bit TIFF image
// Tags: Aspose.Cells TIFF export with LZW compression | ImageOrPrintOptions TiffColorDepth Format8bpp | WorkbookRender multi-page TIFF generation | C# Excel workbook to TIFF conversion | Set TiffColorDepth for 256‑color output

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace WorkbookToTiff8Bit
{
    // The sample loads an Excel workbook, configures ImageOrPrintOptions to use LZW compression and an 8‑bit (256‑color) depth, and employs WorkbookRender to produce a single multi‑page TIFF file. It also handles missing input files, creates the output directory if needed, and reports success or errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Determine input and output file paths (can be overridden via command‑line arguments)
            string inputPath = args.Length > 0 ? args[0] : "input.xlsx";
            string outputPath = args.Length > 1 ? args[1] : "output_8bit.tiff";

            // Verify that the input workbook exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook from the specified file
                Workbook workbook = new Workbook(inputPath);

                // Configure image rendering options for an 8‑bit TIFF
                ImageOrPrintOptions options = new ImageOrPrintOptions
                {
                    ImageType = Aspose.Cells.Drawing.ImageType.Tiff,          // Render to TIFF format
                    TiffCompression = TiffCompression.CompressionLZW,       // Use LZW compression
                    TiffColorDepth = ColorDepth.Format8bpp                  // 8‑bit (256 colors) depth
                };

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Render the entire workbook to a single multi‑page TIFF file
                WorkbookRender renderer = new WorkbookRender(workbook, options);
                renderer.ToImage(outputPath);

                Console.WriteLine($"Workbook successfully converted to 8‑bit TIFF: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during conversion:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
