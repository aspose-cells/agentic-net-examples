// Title: Convert Excel worksheet SVG graphics to high‑resolution PNG images with Aspose.Cells in C#
// Common Searches: C# code to export each worksheet page as a PNG using Aspose.Cells | How to rasterize SVG objects in an Excel file to PNG with Aspose.Cells | Aspose.Cells render Excel sheet to PNG with 300 DPI in .NET | Batch convert Excel worksheets to PNG files in a console application | Save Excel worksheet images as PNG for legacy systems lacking SVG support
// Tags: Aspose.Cells SheetRender to PNG | C# rasterize Excel SVG to PNG | ImageOrPrintOptions high‑dpi PNG export | Excel worksheet page rendering Aspose.Cells | Legacy system PNG conversion from Excel

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;          // For ImageType enum
using Aspose.Cells.Rendering;

namespace SvgToPngRasterizer
{
    // The sample loads an Excel workbook, sets ImageOrPrintOptions for 300 DPI PNG output, and uses SheetRender to rasterize each worksheet (including embedded SVG graphics) page‑by‑page into high‑resolution PNG files saved with safe filenames in a designated output folder.
    class Program
    {
        /// <param name="args">
        /// args[0] - path to the source Excel file (required).
        /// args[1] - output directory for PNG files (optional, defaults to a folder named "output" next to the source file).
        /// </param>
        static void Main(string[] args)
        {
            try
            {
                // Validate input arguments
                if (args.Length == 0 || string.IsNullOrWhiteSpace(args[0]))
                {
                    Console.WriteLine("Usage: SvgToPngRasterizer <input.xlsx> [outputFolder]");
                    return;
                }

                string inputPath = args[0];
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                    return;
                }

                // Determine output folder
                string outputFolder = (args.Length > 1 && !string.IsNullOrWhiteSpace(args[1]))
                    ? args[1]
                    : Path.Combine(Path.GetDirectoryName(inputPath) ?? Environment.CurrentDirectory, "output");

                // Ensure the output directory exists
                Directory.CreateDirectory(outputFolder);

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Console.WriteLine($"Loaded workbook \"{Path.GetFileName(inputPath)}\" with {workbook.Worksheets.Count} worksheet(s).");

                // Configure rendering options for PNG output
                ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
                {
                    ImageType = Aspose.Cells.Drawing.ImageType.Png, // Target format
                    OnePagePerSheet = true,                        // Render each sheet as a single page
                    HorizontalResolution = 300,                    // Higher DPI for better quality
                    VerticalResolution = 300
                };

                // Iterate through each worksheet and render to PNG
                for (int sheetIndex = 0; sheetIndex < workbook.Worksheets.Count; sheetIndex++)
                {
                    try
                    {
                        Worksheet sheet = workbook.Worksheets[sheetIndex];
                        string sheetName = sheet.Name;

                        // Create a SheetRender for the current worksheet
                        SheetRender sheetRender = new SheetRender(sheet, renderOptions);

                        // Render each page of the worksheet
                        for (int pageIndex = 0; pageIndex < sheetRender.PageCount; pageIndex++)
                        {
                            // Build a safe filename: SheetName_PageN.png
                            string safeSheetName = MakeFileSystemSafe(sheetName);
                            string outputFile = Path.Combine(outputFolder, $"{safeSheetName}_Page{pageIndex + 1}.png");

                            // Render the page to the file
                            sheetRender.ToImage(pageIndex, outputFile);
                            Console.WriteLine($"Rendered {outputFile}");
                        }

                        // Release resources used by SheetRender
                        sheetRender.Dispose();
                    }
                    catch (Exception exSheet)
                    {
                        Console.WriteLine($"Error processing worksheet index {sheetIndex}: {exSheet.Message}");
                    }
                }

                Console.WriteLine("Rasterization completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during rasterization:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }

        // Replaces invalid filename characters with underscores
        private static string MakeFileSystemSafe(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }
            return name;
        }
    }
}
