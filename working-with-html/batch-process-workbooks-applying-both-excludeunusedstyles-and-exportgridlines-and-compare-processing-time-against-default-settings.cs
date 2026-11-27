// Title: Batch convert multiple Excel workbooks to HTML with RemoveUnusedStyles and ExportGridLines using Aspose.Cells for .NET and compare save performance
// Common Searches: c# batch convert excel files to html with Aspose.Cells and clean up unused styles | how to enable gridlines when saving a workbook to html using Aspose.Cells .NET | benchmark html save time for default versus optimized options in Aspose.Cells | process all xlsx files in a directory and generate html output with Aspose.Cells | measure performance impact of RemoveUnusedStyles and ExportGridLines in Aspose.Cells
// Tags: batch workbook html conversion Aspose.Cells | unused style cleanup Aspose.Cells | gridlines export option Aspose.Cells | html save performance benchmark Aspose.Cells | c# excel folder processing Aspose.Cells

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatchProcessing
{
    // The C# program scans a given folder for Excel workbooks (xlsx, xls, xlsb, xlsm, csv), creates separate output directories for default and optimized HTML, and for each file performs two saves: one with default HtmlSaveOptions and another after calling RemoveUnusedStyles() and setting ExportGridLines = true. Stopwatch records the elapsed milliseconds for each save, logging the times and any errors, and finally reports the total files processed and the locations of the generated HTML files.
    class Program
    {
        static void Main(string[] args)
        {
            // Determine input and output directories
            string inputFolder = args.Length > 0 ? args[0] : "InputWorkbooks";
            string outputFolderOptimized = Path.Combine("Output", "Optimized");
            string outputFolderDefault = Path.Combine("Output", "Default");

            // Ensure output directories exist
            Directory.CreateDirectory(outputFolderOptimized);
            Directory.CreateDirectory(outputFolderDefault);

            // Verify input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder '{inputFolder}' does not exist.");
                return;
            }

            // Supported workbook extensions
            string[] extensions = new[] { ".xlsx", ".xls", ".xlsb", ".xlsm", ".csv" };

            // Collect all workbook files
            var workbookFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.AllDirectories);
            int processedCount = 0;

            foreach (var filePath in workbookFiles)
            {
                // Skip unsupported files
                if (Array.IndexOf(extensions, Path.GetExtension(filePath).ToLower()) < 0)
                    continue;

                // Ensure the file actually exists (safety check)
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(filePath);
                Console.WriteLine($"Processing file: {filePath}");

                // ---------- Default processing (no style removal, gridlines off) ----------
                try
                {
                    Stopwatch swDefault = Stopwatch.StartNew();

                    // Load workbook
                    Workbook wbDefault = new Workbook(filePath);

                    // Default HTML save options (ExportGridLines defaults to false)
                    HtmlSaveOptions defaultOptions = new HtmlSaveOptions();

                    // Save HTML
                    string defaultOutputPath = Path.Combine(outputFolderDefault, $"{fileNameWithoutExt}_default.html");
                    wbDefault.Save(defaultOutputPath, defaultOptions);

                    swDefault.Stop();
                    Console.WriteLine($"  Default save time: {swDefault.ElapsedMilliseconds} ms");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"  Error during default processing of '{filePath}': {ex.Message}");
                }

                // ---------- Optimized processing (remove unused styles, export gridlines) ----------
                try
                {
                    Stopwatch swOptimized = Stopwatch.StartNew();

                    // Load workbook
                    Workbook wbOptimized = new Workbook(filePath);

                    // Remove unused styles
                    wbOptimized.RemoveUnusedStyles();

                    // HTML save options with gridlines exported
                    HtmlSaveOptions optimizedOptions = new HtmlSaveOptions
                    {
                        ExportGridLines = true
                    };

                    // Save HTML
                    string optimizedOutputPath = Path.Combine(outputFolderOptimized, $"{fileNameWithoutExt}_optimized.html");
                    wbOptimized.Save(optimizedOutputPath, optimizedOptions);

                    swOptimized.Stop();
                    Console.WriteLine($"  Optimized save time: {swOptimized.ElapsedMilliseconds} ms");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"  Error during optimized processing of '{filePath}': {ex.Message}");
                }

                processedCount++;
            }

            Console.WriteLine();
            Console.WriteLine($"Batch processing completed. Total files processed: {processedCount}");
            Console.WriteLine($"Optimized HTML files are located in: {Path.GetFullPath(outputFolderOptimized)}");
            Console.WriteLine($"Default HTML files are located in: {Path.GetFullPath(outputFolderDefault)}");
        }
    }
}
