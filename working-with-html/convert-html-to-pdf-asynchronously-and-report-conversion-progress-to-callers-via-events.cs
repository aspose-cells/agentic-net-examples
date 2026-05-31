using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsConversionDemo
{
    // Event arguments to convey progress information
    public class ProgressEventArgs : EventArgs
    {
        public int PageIndex { get; }
        public int PageCount { get; }

        public ProgressEventArgs(int pageIndex, int pageCount)
        {
            PageIndex = pageIndex;
            PageCount = pageCount;
        }
    }

    // Converter that transforms an HTML file to PDF asynchronously
    public class HtmlToPdfConverter
    {
        // Event raised when a page starts saving
        public event EventHandler<ProgressEventArgs>? PageSavingStarted;

        // Event raised when a page finishes saving
        public event EventHandler<ProgressEventArgs>? PageSavingFinished;

        // Asynchronous conversion method
        public Task ConvertAsync(string htmlFilePath, string pdfFilePath)
        {
            return Task.Run(() =>
            {
                try
                {
                    // Verify that the source HTML file exists
                    if (!File.Exists(htmlFilePath))
                        throw new FileNotFoundException($"HTML file not found: {htmlFilePath}");

                    // Load the HTML file into a workbook using the appropriate constructor
                    var loadOptions = new LoadOptions(LoadFormat.Html);
                    var workbook = new Workbook(htmlFilePath, loadOptions);

                    // Configure PDF save options with a page‑saving callback
                    var pdfSaveOptions = new PdfSaveOptions
                    {
                        PageSavingCallback = new PageSavingCallback(this)
                    };

                    // Save the workbook as PDF; the callback will raise progress events
                    workbook.Save(pdfFilePath, pdfSaveOptions);
                }
                catch (Exception ex)
                {
                    // Log or rethrow as needed; here we rethrow to propagate the error to the caller
                    throw new InvalidOperationException("Failed to convert HTML to PDF.", ex);
                }
            });
        }

        // Internal callback implementation that forwards progress to the public events
        private class PageSavingCallback : IPageSavingCallback
        {
            private readonly HtmlToPdfConverter _parent;

            public PageSavingCallback(HtmlToPdfConverter parent)
            {
                _parent = parent;
            }

            public void PageStartSaving(PageStartSavingArgs args)
            {
                // Raise the start‑saving event
                _parent.PageSavingStarted?.Invoke(_parent,
                    new ProgressEventArgs(args.PageIndex + 1, args.PageCount));
            }

            public void PageEndSaving(PageEndSavingArgs args)
            {
                // Raise the end‑saving event
                _parent.PageSavingFinished?.Invoke(_parent,
                    new ProgressEventArgs(args.PageIndex + 1, args.PageCount));
            }
        }
    }

    // Example usage
    class Program
    {
        static async Task Main()
        {
            var converter = new HtmlToPdfConverter();

            // Subscribe to progress events
            converter.PageSavingStarted += (s, e) =>
                Console.WriteLine($"Starting page {e.PageIndex} of {e.PageCount}...");
            converter.PageSavingFinished += (s, e) =>
                Console.WriteLine($"Finished page {e.PageIndex} of {e.PageCount}.");

            string htmlPath = "sample.html";
            string pdfPath = "output.pdf";

            try
            {
                // Perform the conversion asynchronously
                await converter.ConvertAsync(htmlPath, pdfPath);
                Console.WriteLine("HTML to PDF conversion completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Conversion failed: {ex.Message}");
            }
        }
    }
}