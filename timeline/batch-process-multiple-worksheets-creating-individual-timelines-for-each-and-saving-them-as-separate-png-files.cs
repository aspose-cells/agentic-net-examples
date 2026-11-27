// Title: Generate separate high‑resolution PNG timeline images for each worksheet in an Excel workbook using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export each worksheet as a PNG timeline image | batch convert all Excel sheets to high‑resolution PNG files with transparent background | how to render multiple worksheets to one‑page‑per‑sheet PNG using Aspose.Cells | save Excel workbook worksheets as individual PNG images in C# | Aspose.Cells render worksheet to PNG with 300 dpi and transparent background
// Tags: Aspose.Cells SheetRender PNG export | one page per sheet image rendering | high resolution PNG output with ImageOrPrintOptions | transparent background timeline image generation | filesystem safe worksheet image filenames

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace BatchWorksheetTimeline
{
    // The C# program loads an Excel workbook, iterates through every worksheet, and uses Aspose.Cells' SheetRender with ImageOrPrintOptions (PNG format, OnePagePerSheet, 300 dpi, transparent) to create a timeline‑style image for each sheet. Each image is saved as an individual PNG file with a filesystem‑safe name in a designated output folder, handling multi‑page worksheets when needed.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Validate arguments
                if (args.Length < 1)
                {
                    Console.WriteLine("Error: Input workbook path is required.");
                    Console.WriteLine("Usage: BatchWorksheetTimeline.exe <inputWorkbookPath> [outputDirectory]");
                    return;
                }

                string inputPath = args[0];
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: The file \"{inputPath}\" does not exist.");
                    return;
                }

                // Determine output directory
                string outputDir = args.Length >= 2 ? args[1] : Path.Combine(Path.GetDirectoryName(inputPath) ?? "", "TimelineOutputs");
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Console.WriteLine($"Workbook \"{Path.GetFileName(inputPath)}\" loaded successfully.");
                Console.WriteLine($"Number of worksheets: {workbook.Worksheets.Count}");

                // Iterate through each worksheet
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    Worksheet sheet = workbook.Worksheets[i];
                    string sheetName = sheet.Name;

                    // Configure rendering options
                    ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
                    {
                        ImageType = Aspose.Cells.Drawing.ImageType.Png,
                        // Render the whole sheet on a single page (timeline style)
                        OnePagePerSheet = true,
                        // Optional: set higher resolution for better quality
                        HorizontalResolution = 300,
                        VerticalResolution = 300,
                        // Transparent background can be useful for overlaying timelines
                        Transparent = true
                    };

                    // Create the SheetRender object
                    SheetRender sheetRender = new SheetRender(sheet, renderOptions);

                    // Determine how many pages the sheet occupies (usually 1 because of OnePagePerSheet)
                    int pageCount = sheetRender.PageCount;
                    Console.WriteLine($"Rendering worksheet \"{sheetName}\" ({pageCount} page(s)).");

                    // Render each page to a separate PNG file
                    for (int pageIndex = 0; pageIndex < pageCount; pageIndex++)
                    {
                        // Build output file name: SheetName_PageIndex.png (pageIndex starts at 0)
                        string safeSheetName = MakeFileSystemSafe(sheetName);
                        string outputFileName = pageCount == 1
                            ? $"{safeSheetName}.png"
                            : $"{safeSheetName}_Page{pageIndex + 1}.png";

                        string outputPath = Path.Combine(outputDir, outputFileName);

                        // Render and save
                        sheetRender.ToImage(pageIndex, outputPath);
                        Console.WriteLine($"Saved: {outputPath}");
                    }

                    // Dispose the renderer for the current sheet
                    sheetRender.Dispose();
                }

                Console.WriteLine("All worksheets have been processed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }

        /// <summary>
        /// Replaces characters that are invalid in file names with an underscore.
        /// </summary>
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
