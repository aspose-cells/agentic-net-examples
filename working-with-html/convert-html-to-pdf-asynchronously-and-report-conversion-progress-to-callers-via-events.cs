// Title: Async HTML to PDF conversion with progress events using Aspose.Cells for .NET
// Common Searches: c# async convert html file to pdf with progress callback using Aspose.Cells | how to report conversion progress while converting html to pdf with Aspose.Cells | aspnet core html to pdf conversion async method cancellation token Aspose.Cells | event based progress notification for Aspose.Cells HTML to PDF conversion | load html into workbook and save as pdf asynchronously Aspose.Cells example
// Tags: asynchronous html to pdf conversion Aspose.Cells | progress event handling Aspose.Cells PDF generation | HtmlLoadOptions workbook loading Aspose.Cells | PdfSaveOptions configuration Aspose.Cells | cancellation token support Aspose.Cells conversion

using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;

namespace HtmlToPdfConversion
{
    // Event arguments that carry progress percentage (0‑100)
    // The HtmlToPdfConverter class validates input paths, loads an HTML file into an Aspose.Cells Workbook via HtmlLoadOptions, and saves it as a PDF using PdfSaveOptions inside an async Task.Run call. It raises a ProgressChanged event at 0%, 40% (after loading), and 100% (after saving) and accepts an optional CancellationToken for cooperative cancellation.
    public class ConversionProgressEventArgs : EventArgs
    {
        public int Percentage { get; }

        public ConversionProgressEventArgs(int percentage) => Percentage = percentage;
    }

    // Converter class exposing progress events and an async conversion method
    public class HtmlToPdfConverter
    {
        // Fired whenever the conversion progress changes
        public event EventHandler<ConversionProgressEventArgs>? ProgressChanged;

        // Helper to raise the ProgressChanged event safely
        private void OnProgressChanged(int percent) => ProgressChanged?.Invoke(this, new ConversionProgressEventArgs(percent));

        /// <param name="htmlPath">Full path to the source HTML file.</param>
        /// <param name="pdfPath">Full path where the resulting PDF will be saved.</param>
        /// <param name="cancellationToken">Optional token to cancel the operation.</param>
        /// <returns>A task that completes when the conversion finishes.</returns>
        public async Task ConvertAsync(string htmlPath, string pdfPath, CancellationToken cancellationToken = default)
        {
            // Validate arguments early
            if (string.IsNullOrWhiteSpace(htmlPath))
                throw new ArgumentException("HTML path must be provided.", nameof(htmlPath));
            if (string.IsNullOrWhiteSpace(pdfPath))
                throw new ArgumentException("PDF path must be provided.", nameof(pdfPath));

            // Ensure the source file exists
            if (!File.Exists(htmlPath))
                throw new FileNotFoundException("HTML file not found.", htmlPath);

            // Ensure output directory exists
            string? outputDir = Path.GetDirectoryName(pdfPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Run the conversion on a background thread to avoid blocking the caller
            await Task.Run(() =>
            {
                try
                {
                    // 0% – start
                    OnProgressChanged(0);

                    // -------------------------------------------------
                    // Load HTML into a Workbook (Aspose.Cells can read HTML)
                    // -------------------------------------------------
                    var loadOptions = new HtmlLoadOptions();

                    using var workbook = new Workbook(htmlPath, loadOptions);
                    // 40% – HTML loaded
                    OnProgressChanged(40);

                    // -------------------------------------------------
                    // Prepare PDF save options (optional customizations)
                    // -------------------------------------------------
                    var saveOptions = new PdfSaveOptions
                    {
                        // Example: set other options if needed
                        // Compliance = PdfCompliance.PdfA1b // Removed: not available in current API version
                    };

                    // -------------------------------------------------
                    // Save the workbook as PDF
                    // -------------------------------------------------
                    workbook.Save(pdfPath, saveOptions);
                    // 100% – conversion finished
                    OnProgressChanged(100);
                }
                catch (Exception ex)
                {
                    // Propagate exception after raising final progress (optional)
                    OnProgressChanged(0);
                    throw new InvalidOperationException("Error during HTML to PDF conversion.", ex);
                }
            }, cancellationToken).ConfigureAwait(false);
        }
    }

    // Example usage (for illustration; not part of the library)
    class Program
    {
        static async Task Main()
        {
            var converter = new HtmlToPdfConverter();

            // Subscribe to progress events
            converter.ProgressChanged += (sender, e) =>
            {
                Console.WriteLine($"Conversion progress: {e.Percentage}%");
            };

            string htmlFile = @"C:\Input\sample.html";
            string pdfFile = @"C:\Output\sample.pdf";

            try
            {
                await converter.ConvertAsync(htmlFile, pdfFile);
                Console.WriteLine("Conversion completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Conversion failed: {ex.Message}");
            }
        }
    }
}
