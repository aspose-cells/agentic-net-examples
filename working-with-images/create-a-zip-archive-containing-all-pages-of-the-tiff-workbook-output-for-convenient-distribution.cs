// Title: Generate individual TIFF files for each worksheet page and package them into a ZIP archive using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells render each worksheet page to separate TIFF files | How to create a zip file of multiple TIFF images generated from an Excel workbook in .NET | Export Excel worksheet pages as TIFF and compress them with System.IO.Compression | Aspose.Cells multi-page TIFF export and zip archive example in C# | Save worksheet pagination as TIFF files and bundle into a zip using Aspose.Cells
// Tags: Aspose.Cells worksheet page to TIFF rendering | C# zip multiple TIFF files | System.IO.Compression create zip from directory | Aspose.Cells pagination to multi-page TIFF | temporary folder cleanup after image export

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing; // Required for ImageType, ColorDepth, etc.

namespace AsposeCellsTiffZipDemo
{
    // The sample program creates a workbook with enough data to span several printed pages, renders each page to an individual TIFF file using Aspose.Cells, stores the images in a temporary folder, compresses all TIFF files into a ZIP archive with System.IO.Compression, and finally removes the temporary folder.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths (adjust as needed)
            string outputDirectory = Path.Combine(Directory.GetCurrentDirectory(), "TempTiffPages");
            string zipPath = Path.Combine(Directory.GetCurrentDirectory(), "WorkbookPages.zip");

            try
            {
                // Ensure clean temp folder
                if (Directory.Exists(outputDirectory))
                    Directory.Delete(outputDirectory, true);
                Directory.CreateDirectory(outputDirectory);

                // -----------------------------------------------------------------
                // 1. Create a sample workbook with data that spans multiple pages
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "DemoSheet";

                // Populate enough rows/columns to generate several pages when printed
                for (int row = 0; row < 200; row++)
                {
                    for (int col = 0; col < 20; col++)
                    {
                        sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }

                // Optional: adjust page setup to force multiple pages
                sheet.PageSetup.FitToPagesWide = 1;
                sheet.PageSetup.FitToPagesTall = 0; // let it create as many pages as needed

                // -----------------------------------------------------------------
                // 2. Render each page of the worksheet to an individual TIFF file
                // -----------------------------------------------------------------
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Tiff,
                    TiffCompression = TiffCompression.CompressionLZW,
                    TiffColorDepth = ColorDepth.Format24bpp,
                    HorizontalResolution = 150,
                    VerticalResolution = 150
                };

                SheetRender sheetRender = new SheetRender(sheet, imgOptions);
                int pageCount = sheetRender.PageCount;

                Console.WriteLine($"Worksheet '{sheet.Name}' will be rendered into {pageCount} TIFF page(s).");

                for (int pageIndex = 0; pageIndex < pageCount; pageIndex++)
                {
                    try
                    {
                        string tiffFileName = Path.Combine(outputDirectory,
                            $"{sheet.Name}_Page{pageIndex + 1}.tiff");

                        // Render the specific page to a TIFF file
                        sheetRender.ToImage(pageIndex, tiffFileName);
                        Console.WriteLine($"Rendered page {pageIndex + 1} to '{tiffFileName}'.");
                    }
                    catch (Exception exPage)
                    {
                        Console.Error.WriteLine($"Error rendering page {pageIndex + 1}: {exPage.Message}");
                    }
                }

                // -----------------------------------------------------------------
                // 3. Create a ZIP archive containing all generated TIFF files
                // -----------------------------------------------------------------
                try
                {
                    if (File.Exists(zipPath))
                        File.Delete(zipPath);

                    ZipFile.CreateFromDirectory(outputDirectory, zipPath, CompressionLevel.Optimal, false);
                    Console.WriteLine($"All TIFF pages have been zipped into '{zipPath}'.");
                }
                catch (Exception exZip)
                {
                    Console.Error.WriteLine($"Error creating ZIP archive: {exZip.Message}");
                }

                // -----------------------------------------------------------------
                // 4. Clean up temporary TIFF files (optional)
                // -----------------------------------------------------------------
                try
                {
                    if (Directory.Exists(outputDirectory))
                        Directory.Delete(outputDirectory, true);
                    Console.WriteLine("Temporary TIFF files have been removed.");
                }
                catch (Exception exDel)
                {
                    Console.Error.WriteLine($"Error deleting temporary folder: {exDel.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("An error occurred during processing:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
