// Title: Async Convert WordArt‑Heavy Excel to PDF with Page‑Level Progress (Aspose.Cells C#)
// Description: Loads an .xlsx workbook that contains WordArt and gradient fills, configures PdfSaveOptions (IgnoreError and IPageSavingCallback), and saves it to PDF on a background thread via Task.Run, reporting the start and end of each page rendering.
// Keywords: Aspose.Cells async PDF conversion | C# Excel to PDF WordArt | gradient rendering progress callback | IPageSavingCallback example | PdfSaveOptions IgnoreError | background thread Excel export | per‑page PDF save progress | asynchronous workbook conversion
// Common Searches: asynchronous Excel to PDF conversion Aspose.Cells | how to get page progress when saving PDF with gradients | ignore shape errors during PDF export Aspose.Cells | C# convert WordArt workbook to PDF async | IPageSavingCallback usage example
// Developer Intent: Convert a WordArt‑rich Excel workbook to PDF asynchronously while receiving per‑page rendering progress callbacks.
// Use Cases: Batch‑process large, graphic‑intensive reports in a Windows service without blocking the UI. | Log start and completion times for each PDF page to identify slow‑rendering gradients. | Suppress shape and gradient rendering errors to ensure the conversion finishes without exceptions.
// AI Prompts: Write C# code that uses Aspose.Cells to asynchronously convert an Excel file containing WordArt to PDF and logs page start/end via IPageSavingCallback. | Show how to configure PdfSaveOptions with IgnoreError and attach a custom progress callback for PDF saving in Aspose.Cells. | Explain how to extend the async conversion method to accept a CancellationToken and report progress through IProgress<T>.

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an .xlsx workbook that contains WordArt and gradient fills, configures PdfSaveOptions (IgnoreError and IPageSavingCallback), and saves it to PDF on a background thread via Task.Run, reporting the start and end of each page rendering.
public class GradientPdfConverter
{
    // Callback to report progress of each page being saved.
    private class PageProgressCallback : IPageSavingCallback
    {
        public void PageStartSaving(PageStartSavingArgs args)
        {
            // Report start of page rendering
            Console.WriteLine($"[Progress] Starting to render page {args.PageIndex + 1}");
        }

        public void PageEndSaving(PageEndSavingArgs args)
        {
            // Report end of page rendering
            Console.WriteLine($"[Progress] Finished rendering page {args.PageIndex + 1}");
        }
    }

    /// <param name="sourcePath">Full path of the source .xlsx file.</param>
    /// <param name="destPath">Full path where the resulting PDF will be saved.</param>
    /// <returns>A task representing the asynchronous operation.</returns>
    public static async Task ConvertToPdfAsync(string sourcePath, string destPath)
    {
        try
        {
            // Verify that the source file exists to avoid FileNotFoundException.
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source file not found: {sourcePath}");

            // Load the workbook.
            Workbook workbook = new Workbook(sourcePath);

            // Configure PDF save options.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Hide rendering errors (e.g., shape or gradient issues).
                IgnoreError = true,

                // Attach the progress callback.
                PageSavingCallback = new PageProgressCallback()
            };

            // Perform the save operation on a background thread.
            await Task.Run(() => workbook.Save(destPath, pdfOptions));

            Console.WriteLine("Conversion completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during conversion: {ex.Message}");
            throw;
        }
    }

    // Example usage.
    public static async Task Main()
    {
        try
        {
            string sourceFile = "WordArtSample.xlsx";   // Input workbook containing WordArt/gradients
            string outputPdf = "WordArtSample.pdf";     // Desired PDF output

            await ConvertToPdfAsync(sourceFile, outputPdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
