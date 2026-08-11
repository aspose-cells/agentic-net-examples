// Title: Async Excel‑to‑PDF Conversion with WordArt Support and Page‑Rendering Progress (Aspose.Cells C#)
// Description: Demonstrates how to load an Excel workbook that may contain WordArt, configure PdfSaveOptions with a custom IPageSavingCallback to report page‑start and page‑end events, and save the workbook to PDF on a background thread using Task.Run. The async method can be awaited, keeping UI or service threads responsive while tracking gradient rendering progress.
// Keywords: Aspose.Cells async PDF conversion | C# Excel to PDF WordArt | IPageSavingCallback progress | PdfSaveOptions page callback | Task.Run background conversion | .NET async workbook conversion | gradient rendering progress | Excel WordArt PDF export | page rendering notifications
// Common Searches: async convert Excel with WordArt to PDF Aspose.Cells | track page rendering progress when saving PDF with Aspose.Cells | IPageSavingCallback example C# | run Aspose.Cells PDF conversion on background thread | how to report PDF page save progress Aspose
// Developer Intent: Convert an Excel workbook that contains WordArt to PDF without blocking the caller and receive real‑time notifications for each page rendered.
// Use Cases: Embed the async converter in a WPF or WinForms app to keep the UI responsive while showing per‑page progress for large, WordArt‑rich spreadsheets. | Process bulk Excel‑to‑PDF jobs in a Windows service, logging start/end of each page via the custom callback for audit or monitoring purposes. | Expose the conversion through an ASP.NET Core Web API endpoint that accepts an uploaded workbook and returns the generated PDF, handling errors and progress internally.
// AI Prompts: Generate a xUnit test that verifies ConvertWorkbookToPdfAsync invokes GradientRenderingProgressCallback for every page of a multi‑sheet workbook. | Rewrite GradientRenderingProgressCallback to compute and display overall conversion percentage based on total pages. | Show how to call ConvertWorkbookToPdfAsync from an ASP.NET Core controller and return the PDF as a FileResult. | Create a WPF MVVM example that binds the async conversion task to a progress bar using the page‑saving callbacks. | Provide a PowerShell script that uses the compiled DLL to convert a folder of Excel files with WordArt to PDFs asynchronously.

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to load an Excel workbook that may contain WordArt, configure PdfSaveOptions with a custom IPageSavingCallback to report page‑start and page‑end events, and save the workbook to PDF on a background thread using Task.Run. The async method can be awaited, keeping UI or service threads responsive while tracking gradient rendering progress.
public class WorkbookPdfConverter
{
    // Asynchronous method to convert an Excel workbook (which may contain WordArt) to PDF.
    // It reports progress of each page being rendered via a custom PageSavingCallback.
    public async Task ConvertWorkbookToPdfAsync(string sourceFilePath, string destinationPdfPath)
    {
        // Validate input file existence.
        if (!File.Exists(sourceFilePath))
            throw new FileNotFoundException($"Source file not found: {sourceFilePath}");

        // Run the conversion on a background thread to avoid blocking the caller.
        await Task.Run(() =>
        {
            try
            {
                // Load the workbook from the specified source file.
                Workbook workbook = new Workbook(sourceFilePath);

                // Configure PDF save options.
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Attach a custom callback to receive page‑saving progress notifications.
                    PageSavingCallback = new GradientRenderingProgressCallback()
                };

                // Save the workbook as PDF using the configured options.
                workbook.Save(destinationPdfPath, pdfOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Error] Conversion failed: {ex.Message}");
                throw;
            }
        });
    }

    // Custom implementation of IPageSavingCallback to report rendering progress.
    private class GradientRenderingProgressCallback : IPageSavingCallback
    {
        // Called before a page starts being saved.
        public void PageStartSaving(PageStartSavingArgs args)
        {
            // TotalPages may not be available in some versions; fallback to unknown total.
            Console.WriteLine($"[Progress] Starting to render page {args.PageIndex + 1}.");
        }

        // Called after a page has been saved.
        public void PageEndSaving(PageEndSavingArgs args)
        {
            Console.WriteLine($"[Progress] Finished rendering page {args.PageIndex + 1}.");
            if (args.HasMorePages)
            {
                Console.WriteLine("[Progress] More pages remain to be processed...");
            }
        }
    }
}

// Simple console entry point to demonstrate usage.
public class Program
{
    public static async Task Main(string[] args)
    {
        try
        {
            // Example file paths (adjust as needed).
            string inputPath = "input_with_wordart.xlsx";
            string outputPath = "output.pdf";

            var converter = new WorkbookPdfConverter();
            await converter.ConvertWorkbookToPdfAsync(inputPath, outputPath);

            Console.WriteLine("Conversion completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[Fatal] {ex.Message}");
        }
    }
}
