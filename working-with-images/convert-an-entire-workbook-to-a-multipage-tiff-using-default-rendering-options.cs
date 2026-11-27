// Title: How to convert an entire Excel workbook to a multi‑page TIFF file using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells convert whole workbook to multi page TIFF | Aspose.Cells render all worksheets as a single multi‑page TIFF image | Save Excel file as multi‑page TIFF with default options using Aspose.Cells | WorkbookRender ToImage multi‑page TIFF example in .NET | How to generate multi‑page TIFF from Excel using Aspose.Cells C#
// Tags: Aspose.Cells WorkbookRender to multi‑page TIFF | ImageOrPrintOptions set ImageType Tiff | .NET generate multi‑page TIFF from Excel | default rendering options Aspose.Cells TIFF | C# convert Excel workbook to TIFF image

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace WorkbookToMultiPageTiff
{
    // The sample loads an Excel workbook, sets ImageOrPrintOptions.ImageType to Tiff, creates a WorkbookRender object, and calls ToImage to produce a multi‑page TIFF using Aspose.Cells' default rendering settings, while handling file checks and resource cleanup.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input workbook and output TIFF file.
            // Adjust these paths as needed for your environment.
            string inputFile = "input.xlsx";
            string outputFile = "output.tiff";

            try
            {
                // Verify that the input workbook exists.
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Error: The input file \"{inputFile}\" was not found.");
                    return;
                }

                // Ensure the output directory exists.
                string outputDir = Path.GetDirectoryName(outputFile);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Load the workbook.
                Workbook workbook = new Workbook(inputFile);
                Console.WriteLine($"Loaded workbook \"{inputFile}\" successfully.");

                // Configure image/print options for TIFF output.
                ImageOrPrintOptions options = new ImageOrPrintOptions
                {
                    ImageType = Aspose.Cells.Drawing.ImageType.Tiff
                };

                // Create a WorkbookRender object which will handle the conversion.
                WorkbookRender renderer = null;
                try
                {
                    renderer = new WorkbookRender(workbook, options);
                    // Render the entire workbook to a multi‑page TIFF file.
                    // The ToImage(string) overload creates a multi‑page TIFF when ImageType is Tiff.
                    renderer.ToImage(outputFile);
                }
                finally
                {
                    // Dispose the renderer if it implements IDisposable (newer versions do).
                    if (renderer is IDisposable disposableRenderer)
                    {
                        disposableRenderer.Dispose();
                    }
                }

                Console.WriteLine($"Workbook has been successfully converted to a multi‑page TIFF file: \"{outputFile}\"");
            }
            catch (Exception ex)
            {
                // Output any unexpected errors.
                Console.WriteLine("An error occurred during conversion:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
