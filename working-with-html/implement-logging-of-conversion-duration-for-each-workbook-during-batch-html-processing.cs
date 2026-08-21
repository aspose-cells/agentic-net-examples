// Title: Log conversion time for each workbook in Aspose.Cells batch Excel‑to‑HTML conversion (C#)
// Description: C# sample that scans a folder, converts supported Excel files to HTML with Aspose.Cells, measures each workbook’s conversion duration using Stopwatch, and writes the elapsed seconds to the console while handling missing files and errors.
// Keywords: Aspose.Cells | C# | batch conversion | Excel to HTML | conversion timing | Stopwatch | performance logging | console output | HtmlSaveOptions | workbook processing | .NET
// Common Searches: Aspose.Cells log conversion time per workbook | measure Excel to HTML conversion speed C# | batch convert Excel files to HTML with timing | how to track performance of Aspose.Cells conversion | C# Stopwatch batch HTML export Aspose
// Developer Intent: Add performance measurement to a batch Excel‑to‑HTML conversion using Aspose.Cells and output the duration for each workbook.
// Use Cases: Identify individual workbooks that take unusually long to convert. | Create a simple performance report by aggregating conversion times after the batch run. | Provide diagnostic timing information when a conversion fails. | Integrate conversion timing into CI/CD pipelines for regression monitoring. | Log durations to external monitoring tools or log management systems.
// AI Prompts: Write C# code that wraps workbook.Save with a Stopwatch and writes workbook name and elapsed milliseconds to a CSV file. | Modify the program to accumulate total batch time and display a summary of average, fastest, and slowest conversions. | Create a custom IPageSavingCallback that records page‑level timing when HtmlSaveOptions supports page callbacks. | Generate a PowerShell script that runs the compiled exe, captures console output, and stores it in a structured log file.

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace BatchHtmlConversion
{
    // Optional logger implementing IPageSavingCallback (kept for reference)
    // C# sample that scans a folder, converts supported Excel files to HTML with Aspose.Cells, measures each workbook’s conversion duration using Stopwatch, and writes the elapsed seconds to the console while handling missing files and errors.
    class PageSavingLogger : IPageSavingCallback
    {
        public void PageStartSaving(PageStartSavingArgs args)
        {
            Console.WriteLine($"Start saving page {args.PageIndex + 1} of {args.PageCount}");
        }

        public void PageEndSaving(PageEndSavingArgs args)
        {
            Console.WriteLine($"Finished saving page {args.PageIndex + 1}");
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing source Excel files
            string sourceFolder = @"C:\InputExcel";
            // Folder where HTML files will be saved
            string outputFolder = @"C:\OutputHtml";

            // Ensure output directory exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Get all files in the source folder
            string[] files = Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly);

            foreach (string filePath in files)
            {
                // Filter supported Excel formats
                string ext = Path.GetExtension(filePath).ToLowerInvariant();
                if (ext != ".xls" && ext != ".xlsx" && ext != ".xlsm" && ext != ".xlsb")
                    continue;

                // Verify the file still exists before processing
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                Stopwatch sw = Stopwatch.StartNew();

                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Prepare HTML save options
                    HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

                    // NOTE: HtmlSaveOptions does not expose a PageSavingCallback in older versions.
                    // The logger class is retained for reference if a newer version supports it.

                    // Determine output HTML file name
                    string outputFileName = Path.GetFileNameWithoutExtension(filePath) + ".html";
                    string outputPath = Path.Combine(outputFolder, outputFileName);

                    // Save as HTML
                    workbook.Save(outputPath, htmlOptions);

                    sw.Stop();
                    Console.WriteLine($"Converted '{Path.GetFileName(filePath)}' to HTML in {sw.Elapsed.TotalSeconds:F2} seconds.");
                }
                catch (Exception ex)
                {
                    sw.Stop();
                    Console.WriteLine($"Error converting '{Path.GetFileName(filePath)}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
