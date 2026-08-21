// Title: Log per‑page conversion status during multi‑page TIFF export with Aspose.Cells for .NET
// Description: C# sample that monitors the progress of every worksheet page while converting an Excel workbook to a multi‑page TIFF. It builds a workbook, fills rows to create several printable pages, configures page setup, sets TIFF options, and attaches a custom IPageSavingCallback that writes start, finish and continuation messages to the console. The code also handles output folder creation and exception reporting for the ToTiff call.
// Keywords: Aspose.Cells | C# TIFF conversion | IPageSavingCallback | multi‑page TIFF | SheetRender | page progress logging | .NET Excel to TIFF | TiffCompression LZW | Excel pagination | console logging
// Common Searches: Aspose.Cells log page saving progress TIFF | C# custom IPageSavingCallback for multi‑page TIFF | track conversion status per worksheet page Aspose.Cells | monitor errors during Excel to TIFF rendering .NET | how to use SheetRender ToTiff with callbacks
// Developer Intent: Implement real‑time logging for each worksheet page during TIFF conversion to observe progress and capture failures.
// Use Cases: Display console messages when a page starts and finishes saving to gauge rendering speed of large worksheets. | Detect whether additional pages remain after each save to trigger downstream processing or resource cleanup. | Capture and log exceptions from the ToTiff operation while still providing per‑page status updates.
// AI Prompts: Show how to create an IPageSavingCallback that writes page indices to a log file instead of the console during TIFF export with Aspose.Cells. | Provide a C# snippet that retries a failed page save inside PageEndSavingArgs when converting a workbook to a multi‑page TIFF. | Explain how to cancel the TIFF rendering process from within PageStartSavingArgs based on a custom condition.

using System;
using System.Drawing.Imaging;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsTiffConversion
{
    // Custom callback to log page saving progress
    // C# sample that monitors the progress of every worksheet page while converting an Excel workbook to a multi‑page TIFF. It builds a workbook, fills rows to create several printable pages, configures page setup, sets TIFF options, and attaches a custom IPageSavingCallback that writes start, finish and continuation messages to the console. The code also handles output folder creation and exception reporting for the ToTiff call.
    public class CustomPageSavingCallback : IPageSavingCallback
    {
        // Called before a page is saved
        public void PageStartSaving(PageStartSavingArgs args)
        {
            Console.WriteLine($"Starting to save page {args.PageIndex + 1} of {args.PageCount}");
            // args.IsToOutput defaults to true; you can modify it here if needed
        }

        // Called after a page has been saved
        public void PageEndSaving(PageEndSavingArgs args)
        {
            Console.WriteLine($"Finished saving page {args.PageIndex + 1}");
            if (args.HasMorePages)
            {
                Console.WriteLine("More pages will follow...");
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate data to generate multiple pages when rendered
                for (int row = 0; row < 200; row++)
                {
                    sheet.Cells[row, 0].PutValue($"Row {row + 1}");
                }

                // Configure page setup to allow pagination
                PageSetup pageSetup = sheet.PageSetup;
                pageSetup.PaperSize = PaperSizeType.PaperA4;
                pageSetup.Orientation = PageOrientationType.Portrait;
                pageSetup.FitToPagesWide = 1;   // One page wide
                pageSetup.FitToPagesTall = 0;   // Unlimited pages tall

                // Set image options for TIFF conversion
                ImageOrPrintOptions options = new ImageOrPrintOptions
                {
                    // ImageFormat is not required for TIFF; ToTiff handles format internally
                    TiffCompression = TiffCompression.CompressionLZW,
                    OnePagePerSheet = false   // Allow multiple pages in one TIFF
                };

                // Assign the custom page saving callback
                options.PageSavingCallback = new CustomPageSavingCallback();

                // Create a SheetRender with the worksheet and options
                SheetRender renderer = new SheetRender(sheet, options);

                // Ensure the output directory exists
                string outputPath = "output.tiff";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? string.Empty;
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Render the worksheet to a multi‑page TIFF file
                try
                {
                    renderer.ToTiff(outputPath);
                    Console.WriteLine($"TIFF file generated at: {Path.GetFullPath(outputPath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error during TIFF rendering: {ex.Message}");
                }

                // Optionally save the original workbook for reference
                string workbookPath = "source.xlsx";
                string workbookDir = Path.GetDirectoryName(Path.GetFullPath(workbookPath)) ?? string.Empty;
                if (!string.IsNullOrEmpty(workbookDir) && !Directory.Exists(workbookDir))
                {
                    Directory.CreateDirectory(workbookDir);
                }

                try
                {
                    workbook.Save(workbookPath);
                    Console.WriteLine($"Workbook saved at: {Path.GetFullPath(workbookPath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error saving workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
