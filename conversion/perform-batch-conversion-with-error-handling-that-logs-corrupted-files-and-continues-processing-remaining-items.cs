// Title: C# Batch Excel (.xlsx) to PDF Conversion with Error Logging using Aspose.Cells
// Description: A console utility that scans a folder for .xlsx files, converts each workbook to PDF with Aspose.Cells, enables PdfSaveOptions.IgnoreError to suppress rendering issues, logs any conversion failures, and continues processing the remaining files. Ideal for automated, large‑scale document pipelines.
// Keywords: Aspose.Cells C# example | batch Excel to PDF conversion | ignore rendering errors | PdfSaveOptions.IgnoreError | ConversionUtility | error handling in file conversion | corrupted Excel file logging | command‑line PDF generator | GitHub Aspose.Cells sample | automated document workflow
// Common Searches: convert multiple xlsx files to pdf with Aspose.Cells | c# batch excel to pdf ignore errors | aspocells conversionutility example | skip corrupted Excel files during PDF conversion | how to log conversion failures in Aspose.Cells
// Developer Intent: Convert a set of Excel workbooks to PDF, capture conversion errors, and keep the batch running without interruption.
// Use Cases: Nightly job that turns a folder of financial Excel reports into PDFs while recording any files that fail. | Web service that processes user‑uploaded spreadsheets, returns PDFs, and stores a list of problematic uploads for admin review. | Command‑line tool for migration projects that traverses directory trees, converts every .xlsx to PDF, and writes error details to a log file.
// AI Prompts: Write C# code that uses Aspose.Cells to batch convert .xls and .xlsx files to PDF, writes errors to a log file, and lets the user specify input and output directories. | Show how to configure PdfSaveOptions.IgnoreError in Aspose.Cells to suppress rendering errors during Excel‑to‑PDF conversion. | Extend the provided batch conversion sample with retry logic for transient file‑access exceptions and output a summary report of successes and failures.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace BatchConversionExample
{
    // A console utility that scans a folder for .xlsx files, converts each workbook to PDF with Aspose.Cells, enables PdfSaveOptions.IgnoreError to suppress rendering issues, logs any conversion failures, and continues processing the remaining files. Ideal for automated, large‑scale document pipelines.
    class Program
    {
        static void Main()
        {
            // Folder containing source Excel files
            string sourceFolder = @"C:\InputFiles";
            // Folder where converted PDF files will be saved
            string outputFolder = @"C:\OutputFiles";

            try
            {
                // Ensure the source directory exists
                if (!Directory.Exists(sourceFolder))
                {
                    Console.WriteLine($"Source folder does not exist: {sourceFolder}");
                    return;
                }

                // Ensure the output directory exists
                Directory.CreateDirectory(outputFolder);

                // Get all Excel files in the source folder
                string[] excelFiles = Directory.GetFiles(sourceFolder, "*.xlsx");

                foreach (string sourcePath in excelFiles)
                {
                    // Verify the source file still exists
                    if (!File.Exists(sourcePath))
                    {
                        Console.WriteLine($"File not found (skipped): {sourcePath}");
                        continue;
                    }

                    // Determine the destination PDF file path
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(sourcePath);
                    string destPath = Path.Combine(outputFolder, fileNameWithoutExt + ".pdf");

                    try
                    {
                        // Load options for the source Excel file (optional)
                        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

                        // Save options for PDF conversion with error ignoring enabled
                        PdfSaveOptions saveOptions = new PdfSaveOptions
                        {
                            // Hide rendering errors (e.g., shape, chart, image issues)
                            IgnoreError = true
                        };

                        // Perform the conversion using Aspose.Cells utility method
                        ConversionUtility.Convert(sourcePath, loadOptions, destPath, saveOptions);

                        Console.WriteLine($"Successfully converted: {sourcePath} -> {destPath}");
                    }
                    catch (Exception ex)
                    {
                        // Log the error but continue processing the remaining files
                        Console.WriteLine($"Error converting file '{sourcePath}': {ex.Message}");
                    }
                }

                Console.WriteLine("Batch conversion completed.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
