// Title: Async HTML‑to‑PDF conversion with Aspose.Cells and page‑progress events (C#)
// Description: Demonstrates how to load an HTML file into an Aspose.Cells Workbook, convert it to PDF on a background thread, and report each page's saving progress through a custom IPageSavingCallback that raises events to callers.
// Keywords: Aspose.Cells async HTML to PDF | C# PDF export progress event | IPageSavingCallback example | background PDF conversion .NET | page‑by‑page progress Aspose.Cells | Task.Run PDF generation | HTML workbook to PDF C#
// Common Searches: async HTML to PDF conversion Aspose.Cells C# | how to get page progress while saving PDF with Aspose.Cells | implement IPageSavingCallback for PDF export | C# convert HTML file to PDF without blocking UI | report PDF conversion progress events .NET
// Developer Intent: Convert an HTML document to PDF asynchronously and expose real‑time page‑saving progress via events.
// Use Cases: Generate PDF reports from large HTML templates in a Windows service while updating a UI or log with page numbers. | Run HTML‑to‑PDF conversion in a web API endpoint without tying up request threads and send progress to the client via SignalR. | Process batch HTML files on a server, record each page saved for audit trails, and handle failures per page.
// AI Prompts: Create a unit test that asserts the PageSavingProgress event fires with correct page index and total count during ConvertAsync. | Extend HtmlToPdfConverter to accept a CancellationToken and stop the conversion while still emitting progress events. | Write an ASP.NET Core controller that calls ConvertAsync and streams progress updates to the browser using SignalR.

using System;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsHtmlToPdfAsync
{
    // Event arguments for page progress
    // Demonstrates how to load an HTML file into an Aspose.Cells Workbook, convert it to PDF on a background thread, and report each page's saving progress through a custom IPageSavingCallback that raises events to callers.
    public class PageProgressEventArgs : EventArgs
    {
        public int PageIndex { get; }
        public int PageCount { get; }

        public PageProgressEventArgs(int pageIndex, int pageCount)
        {
            PageIndex = pageIndex;
            PageCount = pageCount;
        }
    }

    // Callback implementation that raises an event for each page start
    public class ProgressPageSavingCallback : IPageSavingCallback
    {
        public event EventHandler<PageProgressEventArgs> PageSavingStarted;

        public void PageStartSaving(PageStartSavingArgs args)
        {
            // Raise progress event (page index is zero‑based, add 1 for human readable)
            PageSavingStarted?.Invoke(this, new PageProgressEventArgs(args.PageIndex + 1, args.PageCount));
        }

        public void PageEndSaving(PageEndSavingArgs args)
        {
            // No additional handling required for this example
        }
    }

    public class HtmlToPdfConverter
    {
        // Event exposed to callers to receive progress updates
        public event EventHandler<PageProgressEventArgs> PageSavingProgress;

        // Asynchronous conversion method
        public Task ConvertAsync(string htmlFilePath, string pdfFilePath)
        {
            return Task.Run(() =>
            {
                // Load the HTML file into a workbook
                Workbook workbook = new Workbook(htmlFilePath);

                // Configure PDF save options with a page‑saving callback
                PdfSaveOptions pdfOptions = new PdfSaveOptions();
                ProgressPageSavingCallback callback = new ProgressPageSavingCallback();

                // Forward callback events to the public event
                callback.PageSavingStarted += (s, e) => PageSavingProgress?.Invoke(this, e);
                pdfOptions.PageSavingCallback = callback;

                // Save the workbook as PDF; the callback will be invoked per page
                workbook.Save(pdfFilePath, pdfOptions);
            });
        }
    }

    class Program
    {
        static async Task Main(string[] args)
        {
            // Example file paths (adjust as needed)
            string htmlPath = "sample.html";
            string pdfPath = "output.pdf";

            HtmlToPdfConverter converter = new HtmlToPdfConverter();

            // Subscribe to progress events
            converter.PageSavingProgress += (sender, e) =>
            {
                Console.WriteLine($"Saving page {e.PageIndex} of {e.PageCount}");
            };

            try
            {
                // Perform conversion asynchronously
                await converter.ConvertAsync(htmlPath, pdfPath);
                Console.WriteLine("HTML to PDF conversion completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Conversion failed: {ex.Message}");
            }
        }
    }
}
