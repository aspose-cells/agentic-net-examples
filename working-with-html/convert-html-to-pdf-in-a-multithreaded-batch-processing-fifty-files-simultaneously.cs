// Title: Convert up to 50 HTML files to PDF concurrently with Aspose.Cells in C# using SemaphoreSlim
// AI Prompts: Write a C# console application that scans a directory for *.html files, loads each into an Aspose.Cells Workbook, and saves them as PDF while executing up to 50 conversions in parallel. | Add configurable parallelism and detailed error logging to the Aspose.Cells HTML‑to‑PDF batch processor. | Implement cancellation support with a CancellationToken so the multi‑threaded conversion can be stopped gracefully.
// Common Searches: c# aspose.cells convert multiple html files to pdf with parallel tasks | how to process 50 html to pdf simultaneously using semaphore in .net | batch html to pdf conversion asp.net core aspose.cells example | limit concurrent file conversions to 50 in C# using SemaphoreSlim | async html to pdf conversion aspose.cells performance tips
// Tags: Aspose.Cells HTML to PDF batch processing | C# SemaphoreSlim parallel file conversion | multi‑threaded HTML to PDF conversion | limit concurrent conversions to 50 | asynchronous workbook loading with HtmlLoadOptions | error handling in parallel Aspose.Cells tasks

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;

// The sample enumerates up to 50 .html files from a source folder, loads each file into an Aspose.Cells Workbook via HtmlLoadOptions, saves it as a .pdf in a target folder, and runs the conversions concurrently. A SemaphoreSlim caps the parallel tasks at 50, ensuring controlled resource usage while providing logging and error handling.
class HtmlToPdfBatchProcessor
{
    // Entry point
    static async Task Main(string[] args)
    {
        try
        {
            // Folder containing HTML files
            string inputFolder = @"C:\InputHtml";
            // Folder where PDF files will be saved
            string outputFolder = @"C:\OutputPdf";

            // Verify input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.Error.WriteLine($"Input folder not found: {inputFolder}");
                return;
            }

            // Ensure output directory exists
            Directory.CreateDirectory(outputFolder);

            // Get all HTML files (adjust the search pattern if needed)
            string[] htmlFiles = Directory.GetFiles(inputFolder, "*.html", SearchOption.TopDirectoryOnly);

            if (htmlFiles.Length == 0)
            {
                Console.WriteLine("No HTML files found to process.");
                return;
            }

            // Limit to 50 files for this batch (or process all if less than 50)
            int batchSize = Math.Min(50, htmlFiles.Length);
            var batchFiles = new List<string>(htmlFiles).GetRange(0, batchSize);

            // Use a semaphore to limit concurrency to 50 simultaneous tasks
            using var semaphore = new SemaphoreSlim(50);

            var tasks = new List<Task>();

            foreach (string htmlPath in batchFiles)
            {
                await semaphore.WaitAsync();

                // Start a task for each file
                tasks.Add(Task.Run(() =>
                {
                    try
                    {
                        // Load HTML into a Workbook
                        var wb = new Workbook(htmlPath, new HtmlLoadOptions());

                        // Determine output PDF path
                        string pdfFileName = Path.GetFileNameWithoutExtension(htmlPath) + ".pdf";
                        string pdfPath = Path.Combine(outputFolder, pdfFileName);

                        // Save as PDF
                        wb.Save(pdfPath, SaveFormat.Pdf);
                    }
                    catch (Exception ex)
                    {
                        // Log or handle errors as needed
                        Console.Error.WriteLine($"Error processing '{htmlPath}': {ex.Message}");
                    }
                    finally
                    {
                        // Release the semaphore slot
                        semaphore.Release();
                    }
                }));
            }

            // Wait for all tasks to complete
            await Task.WhenAll(tasks);

            Console.WriteLine("Batch conversion completed.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.Error.WriteLine($"Fatal error: {ex.Message}");
        }
    }
}
