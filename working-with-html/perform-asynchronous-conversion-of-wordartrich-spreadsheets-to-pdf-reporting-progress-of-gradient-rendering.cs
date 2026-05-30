using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsAsyncPdfConversion
{
    // Callback to report progress while each PDF page is being saved.
    public class PdfPageSavingCallback : IPageSavingCallback
    {
        public void PageStartSaving(PageStartSavingArgs args)
        {
            // TotalPages property is not available in PageStartSavingArgs; report only the page index.
            Console.WriteLine($"[Progress] Starting to render page {args.PageIndex + 1}");
        }

        public void PageEndSaving(PageEndSavingArgs args)
        {
            Console.WriteLine($"[Progress] Finished rendering page {args.PageIndex + 1}");
            if (args.HasMorePages)
            {
                Console.WriteLine("[Progress] More pages pending...");
            }
        }
    }

    public static class WorkbookPdfConverter
    {
        // Asynchronously converts an Excel workbook (including WordArt) to PDF with progress reporting.
        public static async Task ConvertToPdfAsync(string sourceFilePath, string destinationPdfPath)
        {
            try
            {
                // Verify source file existence before loading.
                if (!File.Exists(sourceFilePath))
                {
                    Console.WriteLine($"Source file not found: {sourceFilePath}");
                    return;
                }

                // Load the workbook (Aspose.Cells loads synchronously).
                Workbook workbook = new Workbook(sourceFilePath);

                // Configure PDF save options.
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Suppress rendering errors (e.g., shape or gradient issues).
                    IgnoreError = true,
                    // Attach progress callback.
                    PageSavingCallback = new PdfPageSavingCallback()
                };

                // Save on a background thread to avoid blocking the caller.
                await Task.Run(() => workbook.Save(destinationPdfPath, pdfOptions));

                Console.WriteLine("Conversion completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }

    class Program
    {
        static async Task Main(string[] args)
        {
            try
            {
                string excelPath = "WordArtRichWorkbook.xlsx";
                string pdfPath = "ConvertedOutput.pdf";

                // Ensure the source file exists before starting conversion.
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"Source file not found: {excelPath}");
                    return;
                }

                await WorkbookPdfConverter.ConvertToPdfAsync(excelPath, pdfPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}