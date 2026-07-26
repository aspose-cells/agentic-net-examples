// Title: Convert HTML to PDF in Parallel (up to 50 files) with Aspose.Cells C#
// Description: A C# example that scans a directory for *.html files, loads each into an Aspose.Cells Workbook (LoadFormat.Html) and saves it as PDF (SaveFormat.Pdf). The conversion runs inside Parallel.ForEach with MaxDegreeOfParallelism set to 50, providing fast batch processing and per‑file error logging.
// Keywords: Aspose.Cells | HTML to PDF | C# parallel conversion | batch processing | Parallel.ForEach | max degree of parallelism 50 | multi‑threaded conversion | Workbook LoadFormat.Html | SaveFormat.Pdf | folder batch conversion
// Common Searches: Aspose.Cells convert multiple HTML files to PDF C# | parallel HTML to PDF conversion with Aspose.Cells | batch convert HTML folder to PDF using Parallel.ForEach | limit Aspose.Cells conversion to 50 concurrent tasks | error handling in multi‑threaded HTML to PDF conversion
// Developer Intent: The developer needs to transform every HTML file in a given folder into a PDF document using Aspose.Cells, while processing up to 50 files simultaneously to maximize throughput.
// Use Cases: Nightly job that archives web‑generated reports by converting a large HTML dump to PDF. | Web service that receives bulk HTML spreadsheets and returns PDFs without blocking other requests. | Command‑line tool for finance teams to batch‑convert thousands of HTML invoices to PDF in minutes.
// AI Prompts: Add CancellationToken support to the batch processor so the conversion can be stopped gracefully. | Replace console output with structured logging (e.g., Serilog) while keeping the parallel workflow intact. | Create a unit‑test suite that mocks file I/O, verifies that each HTML file produces a PDF, and checks error handling.

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellsHtmlToPdfBatch
{
    // This class performs multi‑threaded conversion of HTML files to PDF.
    // It processes up to 50 files in parallel using Parallel.ForEach.
    // A C# example that scans a directory for *.html files, loads each into an Aspose.Cells Workbook (LoadFormat.Html) and saves it as PDF (SaveFormat.Pdf). The conversion runs inside Parallel.ForEach with MaxDegreeOfParallelism set to 50, providing fast batch processing and per‑file error logging.
    public static class HtmlToPdfBatchProcessor
    {
        // Converts all *.html files found in inputFolder to PDF files in outputFolder.
        public static void Process(string inputFolder, string outputFolder)
        {
            // Verify that the source directory exists.
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"[Error] Input folder does not exist: {inputFolder}");
                return;
            }

            // Get all HTML files from the source directory.
            string[] htmlFiles = Directory.GetFiles(inputFolder, "*.html", SearchOption.TopDirectoryOnly);
            if (htmlFiles.Length == 0)
            {
                Console.WriteLine($"[Info] No HTML files found in: {inputFolder}");
                return;
            }

            // Ensure the destination directory exists.
            Directory.CreateDirectory(outputFolder);

            // Limit the degree of parallelism to 50 simultaneous tasks.
            ParallelOptions parallelOptions = new ParallelOptions { MaxDegreeOfParallelism = 50 };

            // Process each file concurrently.
            Parallel.ForEach(htmlFiles, parallelOptions, htmlPath =>
            {
                try
                {
                    // Load the HTML file into a Workbook. LoadOptions specifies the source format.
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
                    Workbook workbook = new Workbook(htmlPath, loadOptions);

                    // Build the output PDF file name.
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(htmlPath);
                    string pdfPath = Path.Combine(outputFolder, fileNameWithoutExt + ".pdf");

                    // Save the workbook as PDF.
                    workbook.Save(pdfPath, SaveFormat.Pdf);

                    Console.WriteLine($"[Success] {Path.GetFileName(htmlPath)} -> {Path.GetFileName(pdfPath)}");
                }
                catch (Exception ex)
                {
                    // Log any conversion errors without stopping other tasks.
                    Console.WriteLine($"[Error] Converting '{htmlPath}' failed: {ex.Message}");
                }
            });
        }
    }

    // Example entry point.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Define source and destination folders (adjust as needed).
                string sourceFolder = @"C:\InputHtml";
                string destinationFolder = @"C:\OutputPdf";

                // Run the batch conversion.
                HtmlToPdfBatchProcessor.Process(sourceFolder, destinationFolder);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Fatal] Unexpected error: {ex.Message}");
            }
        }
    }
}
