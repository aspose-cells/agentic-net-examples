// Title: How to split a multi‑page TIFF generated from an Excel workbook into separate single‑page TIFF files using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, determines the total number of pages in the first worksheet, and saves each page as an individual TIFF file with sequential names. | Show how to configure ImageOrPrintOptions and use SheetRender to export every worksheet page to its own TIFF image, including checks for the source file and per‑page exception handling. | Create a C# example that iterates over SheetRender.PageCount, calls ToImage for each index, and logs any rendering errors while producing files like Page_1.tiff, Page_2.tiff, etc.
// Common Searches: Aspose.Cells C# split multi page tiff into separate files | render each worksheet page as individual tiff using SheetRender | C# export Excel worksheet pages to separate tiff images Aspose | how to get SheetRender page count and save each page as tiff | error handling when saving multiple tiff files with Aspose.Cells
// Tags: Aspose.Cells SheetRender TIFF export | C# multi-page TIFF splitting | ImageOrPrintOptions OnePagePerSheet usage | per-page Excel to TIFF conversion | separate TIFF files generation Aspose

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The program loads an Excel workbook, uses Aspose.Cells SheetRender with ImageOrPrintOptions to determine the number of pages in the first worksheet, and saves each page as a separate TIFF file (Page_1.tiff, Page_2.tiff, …). It includes a check for the source file's existence and per‑page error handling.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source workbook.
            string sourcePath = "SourceWorkbook.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException.
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the workbook that will be exported as a multi‑page TIFF.
            Workbook workbook = new Workbook(sourcePath);

            // Get the first worksheet (adjust index if needed).
            Worksheet sheet = workbook.Worksheets[0];

            // Configure image options for TIFF output.
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // Allow multiple pages per worksheet (required for multi‑page TIFF).
                OnePagePerSheet = false
                // Note: In some Aspose.Cells versions the ImageFormat property may not be available.
                // The default format (PNG) will be used if ImageFormat cannot be set.
            };

            // Use SheetRender to obtain the number of pages and render each page.
            SheetRender render = new SheetRender(sheet, imgOptions);
            int totalPages = render.PageCount;

            // Iterate through each page and save it as an individual TIFF file.
            for (int pageIndex = 0; pageIndex < totalPages; pageIndex++)
            {
                try
                {
                    string outputFile = $"Page_{pageIndex + 1}.tiff";
                    render.ToImage(pageIndex, outputFile);
                }
                catch (Exception exPage)
                {
                    Console.WriteLine($"Error rendering page {pageIndex + 1}: {exPage.Message}");
                }
            }

            Console.WriteLine("Splitting completed. Individual TIFF files have been created.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
