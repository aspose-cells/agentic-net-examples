// Title: Generate PNG thumbnails for each page of an Excel workbook using Aspose.Cells in C#
// Common Searches: c# Aspose.Cells render each worksheet page as a PNG thumbnail | how to create a thumbnail gallery from Excel workbook pages with Aspose.Cells | extract individual page images from an Excel file to PNG files using Aspose.Cells | save workbook pages as low‑resolution PNG files for preview in .NET
// Tags: WorkbookRender render Excel pages to PNG | ImageOrPrintOptions low DPI thumbnail generation | Aspose.Cells extract workbook pages as images | C# create thumbnail gallery from Excel workbook | save rendered pages to folder

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;   // Required for ImageType enum

namespace ExtractTiffPagesToThumbnails
{
    // The sample loads an Excel workbook, configures ImageOrPrintOptions for PNG output with reduced DPI, uses WorkbookRender to obtain the page count, iterates through each page, and writes each rendered page as a separate PNG file into a Thumbnails folder, producing a ready‑to‑use thumbnail gallery for preview.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths – adjust as needed
            string workbookPath = "input.xlsx";          // source workbook
            string thumbnailFolder = "Thumbnails";       // folder for thumbnail images

            try
            {
                // ------------------------------------------------------------
                // 1. Verify workbook file exists
                // ------------------------------------------------------------
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Workbook not found: {workbookPath}");
                    return;
                }

                // ------------------------------------------------------------
                // 2. Load the workbook
                // ------------------------------------------------------------
                Workbook workbook = new Workbook(workbookPath);

                // ------------------------------------------------------------
                // 3. Prepare rendering options for PNG thumbnails
                // ------------------------------------------------------------
                ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Png,          // render as PNG (thumbnail)
                    OnePagePerSheet = false,            // render each page separately
                    HorizontalResolution = 150,         // lower DPI for smaller size
                    VerticalResolution = 150
                };

                // ------------------------------------------------------------
                // 4. Ensure output folder exists
                // ------------------------------------------------------------
                if (!Directory.Exists(thumbnailFolder))
                {
                    Directory.CreateDirectory(thumbnailFolder);
                }

                // ------------------------------------------------------------
                // 5. Render each page to a PNG file (thumbnail)
                // ------------------------------------------------------------
                WorkbookRender wbRender = new WorkbookRender(workbook, renderOptions);
                int pageCount = wbRender.PageCount;
                Console.WriteLine($"Workbook contains {pageCount} page(s).");

                for (int i = 0; i < pageCount; i++)
                {
                    try
                    {
                        string thumbPath = Path.Combine(thumbnailFolder, $"thumb_page_{i + 1}.png");
                        wbRender.ToImage(i, thumbPath);   // render page i directly to PNG
                        Console.WriteLine($"Saved thumbnail: {thumbPath}");
                    }
                    catch (Exception pageEx)
                    {
                        Console.WriteLine($"Error processing page {i + 1}: {pageEx.Message}");
                    }
                }

                Console.WriteLine("Thumbnail gallery creation completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
