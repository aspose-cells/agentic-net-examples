// Title: Batch convert 50 HTML files to PDF in parallel with Aspose.Cells (C#)
// Description: Shows how to load 50 HTML documents into Aspose.Cells workbooks and export each to PDF simultaneously using Parallel.ForEach, with automatic output folder creation, per‑file error handling, and optional MaxDegreeOfParallelism tuning.
// Keywords: Aspose.Cells | HTML to PDF | C# parallel conversion | batch PDF generation | Parallel.ForEach | LoadFormat.Html | PdfSaveOptions | .NET | multi‑threaded conversion | GitHub Aspose.Cells example
// Common Searches: convert multiple html files to pdf c# aspose.cells | parallel html to pdf conversion .net | maxdegreeofparallelism aspose.cells pdf export | batch html to pdf aspose example github | error handling parallel file conversion aspose.cells
// Developer Intent: Generate PDF files from a large set of HTML documents concurrently using Aspose.Cells in a C# application.
// Use Cases: Produce PDF reports from HTML templates for dozens of clients in a single run to cut processing time. | Automate conversion of uploaded HTML invoices to PDF in a web service with multi‑threaded throughput. | Archive a directory of marketing HTML assets as PDFs while logging any conversion failures.
// AI Prompts: Add cancellation token support to the parallel HTML‑to‑PDF conversion code using Aspose.Cells. | Show how to set MaxDegreeOfParallelism based on the machine’s CPU core count for optimal performance. | Modify the example to write conversion results and errors to a CSV log file instead of the console.

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

namespace HtmlToPdfBatch
{
    // Shows how to load 50 HTML documents into Aspose.Cells workbooks and export each to PDF simultaneously using Parallel.ForEach, with automatic output folder creation, per‑file error handling, and optional MaxDegreeOfParallelism tuning.
    class Program
    {
        static void Main(string[] args)
        {
            // Prepare a list of HTML files to be converted.
            // In a real scenario these could be read from a directory or a database.
            List<string> htmlFiles = new List<string>();
            for (int i = 1; i <= 50; i++)
            {
                string fileName = $"input_{i}.html";
                htmlFiles.Add(fileName);
            }

            // Ensure that the output directory exists.
            string outputDir = "PdfOutput";
            Directory.CreateDirectory(outputDir);

            // Convert each HTML file to PDF in parallel.
            // ParallelOptions can limit the degree of parallelism if needed.
            ParallelOptions options = new ParallelOptions
            {
                // MaxDegreeOfParallelism = 50; // optional, default is the number of processors
            };

            Parallel.ForEach(htmlFiles, options, htmlPath =>
            {
                try
                {
                    // Verify source file exists.
                    if (!File.Exists(htmlPath))
                    {
                        Console.WriteLine($"Source file not found: {htmlPath}");
                        return;
                    }

                    // Load the HTML file into a workbook.
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
                    Workbook workbook = new Workbook(htmlPath, loadOptions);

                    // Prepare PDF save options (default options are sufficient for most cases).
                    PdfSaveOptions pdfOptions = new PdfSaveOptions();

                    // Determine the output PDF file name.
                    string pdfFileName = Path.GetFileNameWithoutExtension(htmlPath) + ".pdf";
                    string pdfPath = Path.Combine(outputDir, pdfFileName);

                    // Save the workbook as PDF.
                    workbook.Save(pdfPath, pdfOptions);

                    Console.WriteLine($"Converted '{htmlPath}' to '{pdfPath}'.");
                }
                catch (Exception ex)
                {
                    // Log any errors that occur during conversion of an individual file.
                    Console.WriteLine($"Error converting '{htmlPath}': {ex.Message}");
                }
            });

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
