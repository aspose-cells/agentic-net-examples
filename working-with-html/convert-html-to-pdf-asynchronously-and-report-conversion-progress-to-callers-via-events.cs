// Title: Asynchronously convert HTML to PDF with Aspose.Cells in C# and track progress via events
// AI Prompts: Create an async C# method that loads an HTML file into an Aspose.Cells Workbook, saves it as a PDF, and raises a ProgressChanged event at key stages (start, load, save, finish). | Integrate a CancellationToken into the HTML‑to‑PDF conversion routine and adjust the progress‑reporting logic to handle cancellation requests safely. | Demonstrate how to subscribe to the HtmlToPdfConverter.ProgressChanged event and output the percentage and message to the console during conversion.
// Common Searches: asp.net core async html to pdf conversion using Aspose.Cells with progress callback | c# convert html file to pdf with Aspose.Cells and get conversion percentage updates | how to cancel Aspose.Cells html to pdf conversion using CancellationToken | event based progress reporting for Aspose.Cells workbook save to PDF in C#
// Tags: asynchronous html-to-pdf conversion Aspose.Cells | progress event Aspose.Cells workbook save | cancellation token Aspose.Cells conversion | HtmlLoadOptions usage Aspose.Cells | PdfSaveOptions Aspose.Cells PDF export

using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;

namespace HtmlToPdfConversion
{
    // Event arguments for progress reporting
    // Provides an async C# example that loads an HTML document into an Aspose.Cells Workbook, saves it as a PDF, reports conversion progress through a ProgressChanged event at 0%, 10%, 70%, and 100%, and supports cancellation via a CancellationToken.
    public class ProgressChangedEventArgs : EventArgs
    {
        public int Percentage { get; }
        public string Message { get; }

        public ProgressChangedEventArgs(int percentage, string message)
        {
            Percentage = percentage;
            Message = message;
        }
    }

    // Converter class that performs HTML → PDF conversion asynchronously
    public class HtmlToPdfConverter
    {
        // Event raised whenever conversion progress changes
        public event EventHandler<ProgressChangedEventArgs>? ProgressChanged;

        // Helper method to raise the ProgressChanged event safely
        private void OnProgressChanged(int percentage, string message)
        {
            ProgressChanged?.Invoke(this, new ProgressChangedEventArgs(percentage, message));
        }

        /// <param name="htmlPath">Full path to the source HTML file.</param>
        /// <param name="pdfPath">Full path where the resulting PDF will be saved.</param>
        /// <param name="cancellationToken">Optional token to cancel the operation.</param>
        /// <returns>A task that completes when the conversion finishes.</returns>
        public async Task ConvertAsync(string htmlPath, string pdfPath, CancellationToken cancellationToken = default)
        {
            // Validate arguments
            if (string.IsNullOrWhiteSpace(htmlPath))
                throw new ArgumentException("HTML path must be provided.", nameof(htmlPath));
            if (string.IsNullOrWhiteSpace(pdfPath))
                throw new ArgumentException("PDF path must be provided.", nameof(pdfPath));

            // Ensure source HTML file exists
            if (!File.Exists(htmlPath))
                throw new FileNotFoundException("The specified HTML file was not found.", htmlPath);

            // Report start
            OnProgressChanged(0, "Conversion started.");

            await Task.Run(() =>
            {
                // 1. Load HTML into a Workbook
                OnProgressChanged(10, "Loading HTML document.");
                var loadOptions = new HtmlLoadOptions();
                var workbook = new Workbook(htmlPath, loadOptions);

                cancellationToken.ThrowIfCancellationRequested();

                // 2. (Optional) Perform any workbook manipulation here
                // For this example we directly proceed to saving.

                // 3. Save as PDF
                OnProgressChanged(70, "Saving PDF document.");
                var saveOptions = new PdfSaveOptions();
                workbook.Save(pdfPath, saveOptions);

                cancellationToken.ThrowIfCancellationRequested();

                // 4. Report completion
                OnProgressChanged(100, "Conversion completed successfully.");
            }, cancellationToken).ConfigureAwait(false);
        }
    }

    // Entry point for the console application
    public static class Program
    {
        public static async Task Main(string[] args)
        {
            // Example usage:
            // args[0] = path to input HTML file
            // args[1] = path to output PDF file
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: HtmlToPdfConversion <input.html> <output.pdf>");
                return;
            }

            string htmlPath = args[0];
            string pdfPath = args[1];

            var converter = new HtmlToPdfConverter();
            converter.ProgressChanged += (sender, e) =>
            {
                Console.WriteLine($"{e.Percentage}% - {e.Message}");
            };

            try
            {
                await converter.ConvertAsync(htmlPath, pdfPath);
                Console.WriteLine("Conversion finished.");
            }
            catch (OperationCanceledException)
            {
                Console.WriteLine("Conversion was canceled.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during conversion: {ex.Message}");
            }
        }
    }
}
