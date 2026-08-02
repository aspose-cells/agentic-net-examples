// Title: Log conversion time for each workbook in a batch Excel‑to‑HTML conversion with Aspose.Cells for .NET (C#)
// Description: C# sample that scans a folder, loads every Excel workbook with Aspose.Cells, measures the time taken to save each file as HTML using a Stopwatch, and writes the elapsed milliseconds to the console. Includes optional IPageSavingCallback hooks for per‑page start/finish logging and robust error handling.
// Keywords: Aspose.Cells | C# | .NET | Excel to HTML conversion | batch processing | conversion timing | Stopwatch logging | performance measurement | IPageSavingCallback | page saving callback | HTML export | error handling
// Common Searches: how to measure Excel to HTML conversion time Aspose.Cells | batch convert Excel files to HTML with timing logs C# | Aspose.Cells log conversion duration per workbook | use Stopwatch with Aspose.Cells HTML export | page saving callback example Aspose.Cells
// Developer Intent: Add precise timing logs to a batch Excel‑to‑HTML conversion so developers can see how long each workbook takes to process.
// Use Cases: Track performance of large‑scale Excel‑to‑HTML migrations. | Identify slow‑processing workbooks in automated pipelines. | Combine overall conversion timing with optional per‑page callbacks for detailed diagnostics. | Log failures and conversion times for audit trails in CI/CD workflows.
// AI Prompts: Write C# code that batch converts Excel files to HTML with Aspose.Cells and logs the conversion time for each file using Stopwatch. | Show how to attach an IPageSavingCallback to HtmlSaveOptions to log page start and end events during HTML export. | Create a robust example that skips non‑Excel files, handles exceptions, and records elapsed milliseconds for successful conversions.

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsBatchHtml
{
    // Implements IPageSavingCallback to demonstrate page‑level callbacks (optional)
    // C# sample that scans a folder, loads every Excel workbook with Aspose.Cells, measures the time taken to save each file as HTML using a Stopwatch, and writes the elapsed milliseconds to the console. Includes optional IPageSavingCallback hooks for per‑page start/finish logging and robust error handling.
    class PageLoggingCallback : IPageSavingCallback
    {
        public void PageStartSaving(PageStartSavingArgs args)
        {
            // Called when a page starts saving – can be used for detailed logging
            Console.WriteLine($"[Page {args.PageIndex + 1}/{args.PageCount}] start saving.");
        }

        public void PageEndSaving(PageEndSavingArgs args)
        {
            // Called when a page finishes saving
            Console.WriteLine($"[Page {args.PageIndex + 1}/{args.PageCount}] finished saving.");
        }
    }

    class BatchHtmlConverter
    {
        // Processes all Excel files in the input folder and converts each to HTML.
        // Logs the time taken for each workbook conversion.
        public void ConvertFolder(string inputFolder, string outputFolder)
        {
            // Verify input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Ensure output directory exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Get all files in the input folder
            string[] allFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
            foreach (string filePath in allFiles)
            {
                // Filter by known Excel extensions
                string ext = Path.GetExtension(filePath).ToLowerInvariant();
                if (ext != ".xls" && ext != ".xlsx" && ext != ".xlsm" && ext != ".xlsb")
                    continue;

                // Ensure the file actually exists before attempting to load
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found (skipped): {filePath}");
                    continue;
                }

                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(filePath);
                string htmlOutputPath = Path.Combine(outputFolder, fileNameWithoutExt + ".html");

                try
                {
                    // Start timing
                    Stopwatch sw = Stopwatch.StartNew();

                    // Load workbook
                    Workbook wb = new Workbook(filePath);

                    // Prepare HTML save options
                    HtmlSaveOptions saveOptions = new HtmlSaveOptions();

                    // Optional: attach a page‑saving callback if the API supports it
                    // Uncomment the following line if your Aspose.Cells version includes PageSavingCallback
                    // saveOptions.PageSavingCallback = new PageLoggingCallback();

                    // Save as HTML
                    wb.Save(htmlOutputPath, saveOptions);

                    // Stop timing
                    sw.Stop();

                    // Log duration
                    Console.WriteLine($"Converted '{Path.GetFileName(filePath)}' to HTML in {sw.ElapsedMilliseconds} ms.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting '{Path.GetFileName(filePath)}': {ex.Message}");
                }
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Example usage:
            string inputDir = @"C:\InputExcels";
            string outputDir = @"C:\OutputHtml";

            BatchHtmlConverter converter = new BatchHtmlConverter();
            converter.ConvertFolder(inputDir, outputDir);
        }
    }
}
