// Title: C# batch conversion of Excel (.xlsx) to PDF with error handling using Aspose.Cells
// Description: A complete C# example that scans a source directory for *.xlsx files, creates an output folder, and converts each workbook to PDF with Aspose.Cells. The code uses PdfSaveOptions.IgnoreError to suppress rendering issues, catches CellsException for corrupted files, logs every failure to a text file, and continues processing the remaining files.
// Keywords: Aspose.Cells batch conversion | C# Excel to PDF | ignore corrupted Excel files | PdfSaveOptions.IgnoreError | ConversionUtility example | .NET Excel PDF conversion | error logging Aspose.Cells | automated Excel PDF batch
// Common Searches: batch convert xlsx to pdf c# aspnet | aspocells ignoreerror example | skip corrupted excel files aspnet conversion | log excel to pdf conversion errors c# | aspocells conversionutility multiple files
// Developer Intent: Convert a collection of Excel workbooks to PDF, automatically skip files that are damaged, and record any conversion errors without stopping the batch process.
// Use Cases: Nightly processing of uploaded financial reports: convert each Excel file to PDF and capture files that cannot be opened. | Web service that receives user spreadsheets: generate PDFs on‑the‑fly while ignoring and logging corrupted submissions. | Archival workflow for legacy Excel archives: batch‑convert to PDF, ensuring the job continues even if some files are corrupted.
// AI Prompts: Show a C# Aspose.Cells snippet that batch converts .xlsx files to PDF, skips corrupted workbooks, and writes errors to a log. | Explain how PdfSaveOptions.IgnoreError works and how to catch CellsException.FileCorrupted during Excel‑to‑PDF conversion. | Suggest enhancements to add progress reporting and parallel processing to the batch conversion while keeping robust error handling.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace BatchConversionDemo
{
    // A complete C# example that scans a source directory for *.xlsx files, creates an output folder, and converts each workbook to PDF with Aspose.Cells. The code uses PdfSaveOptions.IgnoreError to suppress rendering issues, catches CellsException for corrupted files, logs every failure to a text file, and continues processing the remaining files.
    class Program
    {
        static void Main()
        {
            // Folder containing source Excel files
            string sourceFolder = @"C:\InputFiles";
            // Folder where converted PDFs will be saved
            string outputFolder = @"C:\OutputFiles";

            // Ensure output folder exists
            Directory.CreateDirectory(outputFolder);

            // Verify source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder not found: {sourceFolder}");
                return;
            }

            // Get all Excel files (you can adjust the pattern as needed)
            string[] sourceFiles = Directory.GetFiles(sourceFolder, "*.xlsx");

            foreach (string sourcePath in sourceFiles)
            {
                try
                {
                    // Prepare load options (default loading of XLSX)
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

                    // Prepare save options for PDF and enable error ignoring
                    PdfSaveOptions saveOptions = new PdfSaveOptions
                    {
                        // Hide any rendering errors (shape, image, chart, etc.)
                        IgnoreError = true
                    };

                    // Destination file path with .pdf extension
                    string destPath = Path.Combine(
                        outputFolder,
                        Path.GetFileNameWithoutExtension(sourcePath) + ".pdf");

                    // Perform conversion using Aspose.Cells utility
                    ConversionUtility.Convert(sourcePath, loadOptions, destPath, saveOptions);

                    Console.WriteLine($"Successfully converted: {sourcePath}");
                }
                catch (CellsException cex) when (cex.Code == ExceptionType.FileCorrupted)
                {
                    // Specific handling for corrupted files
                    Console.WriteLine($"Corrupted file skipped: {sourcePath}");
                    LogError(sourcePath, cex);
                }
                catch (Exception ex)
                {
                    // General error handling – log and continue with next file
                    Console.WriteLine($"Error converting {sourcePath}: {ex.Message}");
                    LogError(sourcePath, ex);
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }

        // Simple logger that appends error information to a text file
        private static void LogError(string filePath, Exception ex)
        {
            string logFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ConversionErrors.log");
            string message = $"{DateTime.Now:u} | File: {filePath} | Error: {ex.Message}{Environment.NewLine}";
            File.AppendAllText(logFile, message);
        }
    }
}
