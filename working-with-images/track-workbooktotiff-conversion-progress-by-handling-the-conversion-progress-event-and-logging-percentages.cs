// Title: Monitor Workbook‑to‑TIFF Conversion Progress with IPageSavingCallback in Aspose.Cells for .NET
// Description: C# example that creates a workbook, fills it with rows to produce several pages, sets ImageOrPrintOptions for a multi‑page LZW‑compressed TIFF, and attaches a custom ProgressCallback implementing IPageSavingCallback. The callback logs the percentage of pages saved during the SheetRender.ToTiff operation.
// Keywords: Aspose.Cells | C# | .NET | TIFF conversion | IPageSavingCallback | progress callback | multi‑page TIFF | LZW compression | SheetRender | conversion percentage | image rendering
// Common Searches: Aspose.Cells track TIFF conversion progress C# | IPageSavingCallback example for multi‑page TIFF | log percentage while exporting workbook to TIFF | monitor Aspose.Cells image rendering progress | C# workbook to TIFF with progress callback
// Developer Intent: The developer needs to observe and log the completion percentage while converting a large workbook to a multi‑page TIFF file.
// Use Cases: Display real‑time conversion status in a console or UI for large Excel exports. | Write page‑by‑page progress to a log file for batch processing audits. | Drive a progress bar in WinForms, WPF, or web applications during TIFF generation. | Integrate conversion metrics into monitoring dashboards for automated workflows.
// AI Prompts: Generate C# code that updates a WinForms ProgressBar using IPageSavingCallback during a multi‑page TIFF export with Aspose.Cells. | Show how to write page‑saving percentages to a CSV file instead of the console in the ProgressCallback example. | Explain how to compute overall progress when converting multiple worksheets with OnePagePerSheet set to true.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsConversionProgressDemo
{
    // Custom callback to track page saving progress
    // C# example that creates a workbook, fills it with rows to produce several pages, sets ImageOrPrintOptions for a multi‑page LZW‑compressed TIFF, and attaches a custom ProgressCallback implementing IPageSavingCallback. The callback logs the percentage of pages saved during the SheetRender.ToTiff operation.
    public class ProgressCallback : IPageSavingCallback
    {
        // Called when a page starts saving
        public void PageStartSaving(PageStartSavingArgs args)
        {
            // Calculate percentage based on current page index and total page count
            int currentPage = args.PageIndex + 1; // pages are zero‑based
            int totalPages = args.PageCount;
            double percent = (double)currentPage / totalPages * 100;

            Console.WriteLine($"Saving page {currentPage}/{totalPages} ({percent:F1}% complete)");
        }

        // Called when a page finishes saving
        public void PageEndSaving(PageEndSavingArgs args)
        {
            // Optionally log when a page is finished
            Console.WriteLine($"Finished page {args.PageIndex + 1}");
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with enough rows to generate multiple pages
            for (int row = 0; row < 300; row++)
            {
                sheet.Cells[row, 0].PutValue($"Row {row + 1}");
            }

            // Configure image options for TIFF conversion
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = ImageType.Tiff,
                TiffCompression = TiffCompression.CompressionLZW,
                OnePagePerSheet = false, // generate multi‑page TIFF
                PageSavingCallback = new ProgressCallback() // attach progress callback
            };

            // Create a SheetRender object with the worksheet and options
            SheetRender renderer = new SheetRender(sheet, options);

            // Render the worksheet to a multi‑page TIFF file
            string outputPath = "WorkbookToTiff_WithProgress.tiff";
            renderer.ToTiff(outputPath);

            Console.WriteLine($"Conversion completed. TIFF saved to: {outputPath}");
        }
    }
}
