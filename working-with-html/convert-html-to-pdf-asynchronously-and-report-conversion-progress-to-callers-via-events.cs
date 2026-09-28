// Title: Convert HTML generated from an Excel workbook to PDF asynchronously with Aspose.Cells and report progress via a ProgressChanged event (C#)
// AI Prompts: Write a C# async method that loads an HTML file into an Aspose.Cells Workbook using HtmlLoadOptions on a background thread, saves it as PDF with PdfSaveOptions on another background thread, and raises a ProgressChanged event at 0%, 50%, and 100%. | Show how to subscribe to the ProgressChanged event from the Aspose.Cells HTML‑to‑PDF converter and display the conversion percentage in the console. | Add error handling that checks for a missing HTML file, creates the output directory if needed, and propagates exceptions during the asynchronous conversion.
// Common Searches: how to convert html to pdf asynchronously using aspose.cells in c# | asp.net core html to pdf conversion with progress event Aspose.Cells | c# load html file into workbook and export to pdf on background thread | report conversion percentage during aspose.cells html to pdf conversion | async method for html to pdf conversion with Aspose.Cells PdfSaveOptions
// Tags: asynchronous html to pdf conversion Aspose.Cells | HtmlLoadOptions workbook initialization | PdfSaveOptions pdf generation | event based conversion progress reporting | background thread workbook processing | input file validation Aspose.Cells

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

namespace HtmlToPdfConversion
{
    // Provides an async ConvertAsync method that validates the HTML input, ensures the output folder exists, loads the HTML into an Aspose.Cells Workbook on a background thread using HtmlLoadOptions, saves it as PDF with PdfSaveOptions on another background thread, and raises a ProgressChanged event at 0%, 50%, and 100% to report conversion progress.
    public class HtmlToPdfConverter
    {
        // Event raised to report conversion progress (percentage 0-100)
        public event EventHandler<int>? ProgressChanged;

        // Helper to raise the ProgressChanged event
        protected virtual void OnProgressChanged(int percent)
        {
            ProgressChanged?.Invoke(this, percent);
        }

        // Asynchronously converts an HTML file (generated from a spreadsheet) to PDF
        public async Task ConvertAsync(string htmlFilePath, string pdfFilePath)
        {
            // Validate input file
            if (!File.Exists(htmlFilePath))
                throw new FileNotFoundException("HTML file not found.", htmlFilePath);

            // Ensure output directory exists
            var outDir = Path.GetDirectoryName(pdfFilePath);
            if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
                Directory.CreateDirectory(outDir);

            // Initial progress
            OnProgressChanged(0);

            Workbook workbook = null;

            // Load the HTML into a Workbook on a background thread
            await Task.Run(() =>
            {
                var loadOptions = new HtmlLoadOptions();
                workbook = new Workbook(htmlFilePath, loadOptions);
            });

            // Report midway progress after loading is complete
            OnProgressChanged(50);

            // Save the workbook as PDF on a background thread
            await Task.Run(() =>
            {
                var saveOptions = new PdfSaveOptions();
                workbook.Save(pdfFilePath, saveOptions);
            });

            // Final progress notification
            OnProgressChanged(100);
        }
    }

    class Program
    {
        // Async entry point
        static async Task Main(string[] args)
        {
            // Determine file paths (use defaults if not provided)
            string htmlPath = args.Length > 0 ? args[0] : "input.html";
            string pdfPath = args.Length > 1 ? args[1] : "output.pdf";

            var converter = new HtmlToPdfConverter();
            converter.ProgressChanged += (s, p) => Console.WriteLine($"Progress: {p}%");

            try
            {
                await converter.ConvertAsync(htmlPath, pdfPath);
                Console.WriteLine("Conversion completed successfully.");
            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine($"File error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
