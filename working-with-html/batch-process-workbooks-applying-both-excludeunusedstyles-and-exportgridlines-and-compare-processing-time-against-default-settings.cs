// Title: C# batch conversion of Excel workbooks to HTML with ExcludeUnusedStyles, ExportGridLines and performance comparison
// Description: Scans a directory of .xlsx files, saves each workbook to HTML twice—first with default settings, then with HtmlSaveOptions (ExcludeUnusedStyles = true, ExportGridLines = true)—while measuring and reporting the elapsed time for both approaches.
// Keywords: Aspose.Cells | C# HTML export | ExcludeUnusedStyles | ExportGridLines | batch workbook conversion | performance benchmark | .NET Excel to HTML | SaveFormat.Html | bulk Excel processing | conversion speed comparison
// Common Searches: Aspose.Cells batch export Excel to HTML C# | ExcludeUnusedStyles HtmlSaveOptions example | ExportGridLines performance Aspose.Cells | measure HTML save time for multiple workbooks | compare default and custom HTML export speed .NET
// Developer Intent: The developer wants to convert many Excel files to HTML with specific options and see how those options affect conversion time compared with the default export.
// Use Cases: Generate quick HTML previews of a large Excel archive using default settings. | Create lightweight web pages that include grid lines and omit unused CSS styles. | Benchmark the impact of ExcludeUnusedStyles and ExportGridLines on conversion throughput.
// AI Prompts: Write a C# method that accepts input and output folder paths and returns a dictionary mapping each file name to its default and custom conversion durations. | Show how to log the timing results to a CSV file instead of the console while preserving the batch workflow. | Explain how to safely parallelize the processing loop with Aspose.Cells to improve overall conversion speed.

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// Scans a directory of .xlsx files, saves each workbook to HTML twice—first with default settings, then with HtmlSaveOptions (ExcludeUnusedStyles = true, ExportGridLines = true)—while measuring and reporting the elapsed time for both approaches.
class BatchWorkbookProcessor
{
    static void Main()
    {
        // Folder containing source Excel files
        string sourceFolder = @"C:\Workbooks\Input";
        // Output folders for default and custom HTML saves
        string defaultOutputFolder = @"C:\Workbooks\Output\Default";
        string customOutputFolder = @"C:\Workbooks\Output\Custom";

        // Ensure output directories exist
        Directory.CreateDirectory(defaultOutputFolder);
        Directory.CreateDirectory(customOutputFolder);

        // Verify source folder exists
        if (!Directory.Exists(sourceFolder))
        {
            Console.WriteLine($"Source folder not found: {sourceFolder}");
            return;
        }

        // Process each .xlsx file in the source folder
        foreach (string filePath in Directory.GetFiles(sourceFolder, "*.xlsx"))
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found (skipped): {filePath}");
                    continue;
                }

                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(filePath);

                // ---------- Default save (no explicit options) ----------
                Stopwatch defaultTimer = Stopwatch.StartNew();

                // Load workbook using the standard constructor
                Workbook defaultWb = new Workbook(filePath);
                // Save to HTML using default options
                string defaultHtmlPath = Path.Combine(defaultOutputFolder, fileNameWithoutExt + ".html");
                defaultWb.Save(defaultHtmlPath, SaveFormat.Html);

                defaultTimer.Stop();
                long defaultElapsedMs = defaultTimer.ElapsedMilliseconds;

                // ---------- Custom save (ExcludeUnusedStyles + ExportGridLines) ----------
                Stopwatch customTimer = Stopwatch.StartNew();

                // Load workbook again for a fair comparison
                Workbook customWb = new Workbook(filePath);
                // Create HtmlSaveOptions and set required properties
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExcludeUnusedStyles = true,   // explicitly exclude unused styles
                    ExportGridLines = true        // export grid lines
                };
                // Save to HTML with the custom options
                string customHtmlPath = Path.Combine(customOutputFolder, fileNameWithoutExt + "_grid.html");
                customWb.Save(customHtmlPath, htmlOptions);

                customTimer.Stop();
                long customElapsedMs = customTimer.ElapsedMilliseconds;

                // Output timing comparison
                Console.WriteLine($"File: {Path.GetFileName(filePath)}");
                Console.WriteLine($"  Default save time: {defaultElapsedMs} ms");
                Console.WriteLine($"  Custom save time : {customElapsedMs} ms");
                Console.WriteLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch processing completed.");
    }
}
