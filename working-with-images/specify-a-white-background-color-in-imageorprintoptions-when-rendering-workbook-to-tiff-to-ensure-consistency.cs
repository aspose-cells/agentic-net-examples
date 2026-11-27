// Title: Set a white background for TIFF export of an Excel workbook using Aspose.Cells ImageOrPrintOptions in C#
// Common Searches: c# how to export Excel to TIFF with a solid white background using Aspose.Cells | Aspose.Cells ImageOrPrintOptions Transparent false for TIFF conversion | render multi‑page TIFF from workbook with white background Aspose.Cells C# | set DPI and compression for TIFF output in Aspose.Cells ImageOrPrintOptions | disable transparency when saving Excel as TIFF with Aspose.Cells
// Tags: ImageOrPrintOptions background color TIFF | Aspose.Cells disable transparency TIFF | Aspose.Cells multi-page TIFF rendering C# | Aspose.Cells TIFF LZW compression | Aspose.Cells TIFF 300 dpi resolution

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsTiffWhiteBackground
{
    // The example loads an Excel workbook, configures ImageOrPrintOptions with Transparent set to false (producing a white background), 300 dpi resolution, LZW compression, and 24‑bit color depth, then renders the entire workbook to a multi‑page TIFF file using Aspose.Cells in C#.
    class Program
    {
        static void Main(string[] args)
        {
            // Determine input and output file paths.
            string inputPath = args.Length > 0 ? args[0] : "input.xlsx";
            string outputPath = args.Length > 1 ? args[1] : "output.tiff";

            try
            {
                // Verify that the input file exists before attempting to load it.
                if (!File.Exists(inputPath))
                {
                    Console.Error.WriteLine($"Input file \"{inputPath}\" not found.");
                    Environment.Exit(1);
                }

                // Load the workbook.
                Workbook workbook = new Workbook(inputPath);
                Console.WriteLine($"Loaded workbook \"{inputPath}\" successfully.");

                // Configure rendering options for TIFF output.
                ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
                {
                    // Set the image type to TIFF.
                    ImageType = ImageType.Tiff,

                    // Ensure a white background by disabling transparency.
                    Transparent = false,

                    // Set resolution for higher quality output.
                    HorizontalResolution = 300,
                    VerticalResolution = 300,

                    // Choose a compression method and color depth.
                    TiffCompression = TiffCompression.CompressionLZW,
                    TiffColorDepth = ColorDepth.Format24bpp
                };

                // Render the entire workbook to a multi‑page TIFF file.
                WorkbookRender renderer = new WorkbookRender(workbook, renderOptions);
                renderer.ToImage(outputPath);
                Console.WriteLine($"Workbook rendered to TIFF with white background at \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Report any errors without swallowing them.
                Console.Error.WriteLine("An error occurred during processing:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
                Environment.Exit(1);
            }
        }
    }
}
