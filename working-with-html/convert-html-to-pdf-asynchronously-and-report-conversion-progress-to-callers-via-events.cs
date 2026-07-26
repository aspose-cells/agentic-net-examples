// Title: Async HTML‑to‑PDF Conversion with Page‑Level Progress Using Aspose.Cells for .NET
// Description: Shows how to load an HTML workbook, convert it to PDF on a background thread, and emit start/end messages for each page through a custom IPageSavingCallback that writes to an IProgress<string> instance.
// Keywords: Aspose.Cells | C# | async HTML to PDF | IPageSavingCallback | PdfSaveOptions | progress reporting | Task.Run | LoadOptions Html | .NET PDF conversion
// Common Searches: async html to pdf Aspose.Cells C# | page saving callback Aspose.Cells example | report pdf conversion progress .NET | convert html workbook to pdf asynchronously | IProgress usage with Aspose.Cells
// Developer Intent: Create an asynchronous HTML‑to‑PDF conversion routine that provides page‑by‑page progress notifications.
// Use Cases: Run large HTML report conversions on a background thread and update a UI progress bar with per‑page status. | Expose a Web API endpoint that triggers the conversion and streams progress events to the client in real time. | Log the start and completion time of each PDF page during batch processing for audit or performance analysis.
// AI Prompts: Generate a unit test that verifies HtmlToPdfConverter.ConvertAsync sends a start and end message for every page via IProgress<string>. | Show how to add a CancellationToken to ConvertAsync so the conversion can be cancelled gracefully. | Provide a WPF sample that binds a ProgressBar to the IProgress<string> output of the asynchronous conversion method.

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsHtmlToPdfAsync
{
    // Implements page saving callback to report progress.
    // Shows how to load an HTML workbook, convert it to PDF on a background thread, and emit start/end messages for each page through a custom IPageSavingCallback that writes to an IProgress<string> instance.
    public class ProgressPageSavingCallback : IPageSavingCallback
    {
        private readonly IProgress<string> _progress;

        public ProgressPageSavingCallback(IProgress<string> progress)
        {
            _progress = progress;
        }

        public void PageStartSaving(PageStartSavingArgs args)
        {
            // Report start of a page.
            _progress?.Report($"Starting to save page {args.PageIndex + 1} of {args.PageCount}");
        }

        public void PageEndSaving(PageEndSavingArgs args)
        {
            // Report end of a page.
            _progress?.Report($"Finished saving page {args.PageIndex + 1} of {args.PageCount}");
        }
    }

    public static class HtmlToPdfConverter
    {
        // Asynchronously converts an HTML file to PDF and reports progress.
        public static async Task ConvertAsync(string htmlFilePath, string pdfFilePath, IProgress<string> progress = null)
        {
            // Validate input paths.
            if (string.IsNullOrWhiteSpace(htmlFilePath))
                throw new ArgumentException("HTML file path must be provided.", nameof(htmlFilePath));
            if (string.IsNullOrWhiteSpace(pdfFilePath))
                throw new ArgumentException("PDF file path must be provided.", nameof(pdfFilePath));
            if (!File.Exists(htmlFilePath))
                throw new FileNotFoundException("HTML source file not found.", htmlFilePath);

            await Task.Run(() =>
            {
                // Load the HTML workbook.
                var loadOptions = new LoadOptions(LoadFormat.Html);
                var workbook = new Workbook(htmlFilePath, loadOptions);

                // Configure PDF save options with progress callback.
                var pdfSaveOptions = new PdfSaveOptions
                {
                    PageSavingCallback = new ProgressPageSavingCallback(progress)
                };

                // Save as PDF.
                workbook.Save(pdfFilePath, pdfSaveOptions);
            });
        }
    }

    // Example usage.
    class Program
    {
        static async Task Main(string[] args)
        {
            // Paths to source HTML and destination PDF.
            string htmlPath = "sample.html";
            string pdfPath = "output.pdf";

            // Simple progress reporter that writes to console.
            var progress = new Progress<string>(message => Console.WriteLine(message));

            try
            {
                Console.WriteLine("Conversion started...");
                await HtmlToPdfConverter.ConvertAsync(htmlPath, pdfPath, progress);
                Console.WriteLine("Conversion completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }
}
