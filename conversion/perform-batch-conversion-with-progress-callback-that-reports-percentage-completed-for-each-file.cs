// Title: Batch convert Excel to PDF with per‑page and overall progress using Aspose.Cells for .NET
// Description: Demonstrates how to convert a collection of Excel workbooks to PDF with Aspose.Cells, using an IPageSavingCallback to log each page and a batch routine that reports cumulative percentage, validates files, creates target folders, and handles conversion errors.
// Keywords: Aspose.Cells batch conversion | C# Excel to PDF | IPageSavingCallback | ConversionUtility progress | PdfSaveOptions callback | multiple file conversion .NET | conversion error handling | per‑page progress reporting | overall batch percentage | Aspose.Cells PDF export
// Common Searches: convert multiple Excel files to PDF with Aspose.Cells and show progress | C# Aspose.Cells page saving callback example | batch conversion percentage completed Aspose.Cells .NET | how to handle missing source files in Aspose.Cells batch conversion | Aspose.Cells convert XLSX to PDF with progress callback
// Developer Intent: Convert several Excel workbooks to PDF while displaying both per‑page and total batch progress.
// Use Cases: Automate archival of a folder of .xlsx reports to PDF, logging each page and overall completion. | Provide real‑time feedback in a UI during large spreadsheet PDF generation by attaching a custom IPageSavingCallback. | Run unattended batch jobs that skip absent files, create destination directories, and continue processing after individual conversion failures.
// AI Prompts: Generate a C# method that uses Aspose.Cells ConversionUtility to convert an array of Excel files to PDF with an IPageSavingCallback that logs each page and reports batch percentage. | Show error‑handling code for a batch conversion that validates source files, creates output folders, and proceeds when a file fails. | Explain how to extend the BatchConverter to support PNG or DOCX output while preserving progress reporting.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Utility;

namespace AsposeCellsBatchConversion
{
    // Callback to report progress of page saving during conversion
    // Demonstrates how to convert a collection of Excel workbooks to PDF with Aspose.Cells, using an IPageSavingCallback to log each page and a batch routine that reports cumulative percentage, validates files, creates target folders, and handles conversion errors.
    public class PageProgressCallback : IPageSavingCallback
    {
        public void PageStartSaving(PageStartSavingArgs args)
        {
            // Report start of a page
            Console.WriteLine($"   Saving page {args.PageIndex + 1} of {args.PageCount}");
        }

        public void PageEndSaving(PageEndSavingArgs args)
        {
            // Optionally report end of a page
            // Console.WriteLine($"   Finished page {args.PageIndex + 1}");
        }
    }

    public class BatchConverter
    {
        // Performs batch conversion and reports percentage completed for each file
        public void ConvertFiles(string[] sourceFiles, string[] destinationFiles)
        {
            if (sourceFiles == null) throw new ArgumentNullException(nameof(sourceFiles));
            if (destinationFiles == null) throw new ArgumentNullException(nameof(destinationFiles));
            if (sourceFiles.Length != destinationFiles.Length)
                throw new ArgumentException("Source and destination arrays must have the same length.");

            int total = sourceFiles.Length;
            int processed = 0;

            for (int i = 0; i < total; i++)
            {
                string src = sourceFiles[i];
                string dst = destinationFiles[i];

                // Verify source file exists
                if (!File.Exists(src))
                {
                    Console.WriteLine($"Source file not found: '{src}'. Skipping.");
                    continue;
                }

                // Ensure destination directory exists
                try
                {
                    string destDir = Path.GetDirectoryName(dst);
                    if (!string.IsNullOrEmpty(destDir) && !Directory.Exists(destDir))
                    {
                        Directory.CreateDirectory(destDir);
                    }
                }
                catch (Exception dirEx)
                {
                    Console.WriteLine($"Failed to prepare destination directory for '{dst}': {dirEx.Message}");
                    continue;
                }

                // Create save options with page saving callback
                PdfSaveOptions saveOptions = new PdfSaveOptions
                {
                    PageSavingCallback = new PageProgressCallback()
                };

                // Perform conversion inside a try‑catch to handle runtime errors
                try
                {
                    ConversionUtility.Convert(src, new LoadOptions(), dst, saveOptions);
                    processed++;
                    double percent = (processed * 100.0) / total;
                    Console.WriteLine($"Batch progress: {percent:0.##}% ({processed}/{total}) - Converted '{src}' to '{dst}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting '{src}' to '{dst}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Example file lists (replace with actual paths)
                string[] sources = { "file1.xlsx", "file2.xlsx", "file3.xlsx" };
                string[] destinations = { "file1.pdf", "file2.pdf", "file3.pdf" };

                var converter = new BatchConverter();
                converter.ConvertFiles(sources, destinations);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
