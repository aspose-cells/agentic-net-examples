// Title: Batch add a disclaimer comment to cell A1 of every worksheet in multiple Excel workbooks using Aspose.Cells for .NET
// AI Prompts: Write a C# routine that scans a directory for all supported Excel files and inserts a given disclaimer as a comment in cell A1 of each worksheet using Aspose.Cells. | Implement logic to save the updated workbooks either by overwriting the originals or by writing them to a separate output folder while preserving file names. | Add comprehensive error handling that skips non‑Excel files, logs missing files, and gracefully bypasses password‑protected or corrupted workbooks.
// Common Searches: how to add the same comment to every sheet in a batch of Excel files with Aspose.Cells C# | C# process folder of .xlsx files and insert disclaimer comment in A1 using Aspose.Cells | Aspose.Cells bulk add worksheet note and handle password protected workbooks .NET
// Tags: add worksheet comment Aspose.Cells C# | batch process Excel files Aspose.Cells | insert disclaimer note cell A1 Aspose.Cells | save modified workbooks to separate folder Aspose.Cells | handle password protected Excel Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace BatchWorkbookProcessor
{
    // Iterates through all supported Excel files in a specified folder, adds a provided disclaimer as a comment to cell A1 of each worksheet, and saves the workbooks either in place or to an output directory, while gracefully handling non‑Excel files, missing files, and password‑protected or corrupted workbooks.
    public static class DisclaimerAdder
    {
        /// <param name="inputFolder">Folder containing the workbooks to process.</param>
        /// <param name="disclaimer">The disclaimer text to add as a comment.</param>
        /// <param name="outputFolder">
        /// Optional folder to save the modified workbooks.
        /// If null or empty, the original files are overwritten.
        /// </param>
        public static void ProcessFolder(string inputFolder, string disclaimer, string outputFolder = null)
        {
            if (string.IsNullOrWhiteSpace(inputFolder))
                throw new ArgumentException("Input folder path must be provided.", nameof(inputFolder));

            if (!Directory.Exists(inputFolder))
                throw new DirectoryNotFoundException($"The folder '{inputFolder}' does not exist.");

            // Determine whether to overwrite or write to a separate folder
            bool overwrite = string.IsNullOrWhiteSpace(outputFolder);
            if (!overwrite && !Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Supported Excel extensions
            string[] extensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb", ".ods", ".csv" };

            foreach (string filePath in Directory.EnumerateFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly))
            {
                if (Array.IndexOf(extensions, Path.GetExtension(filePath).ToLowerInvariant()) < 0)
                    continue; // Skip non‑Excel files

                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook; if the file is password‑protected, an exception will be thrown and caught below
                    LoadOptions loadOptions = new LoadOptions();
                    using (Workbook workbook = new Workbook(filePath, loadOptions))
                    {
                        // Add disclaimer comment to cell A1 of each worksheet
                        foreach (Worksheet sheet in workbook.Worksheets)
                        {
                            int commentIndex = sheet.Comments.Add("A1");
                            sheet.Comments[commentIndex].Note = disclaimer;
                            sheet.Comments[commentIndex].Author = "System";
                        }

                        // Determine the save path
                        string savePath = overwrite
                            ? filePath
                            : Path.Combine(outputFolder, Path.GetFileName(filePath));

                        // Save the modified workbook
                        workbook.Save(savePath);
                    }
                }
                catch (CellsException ex)
                {
                    // Handle password‑protected or corrupted files gracefully
                    Console.WriteLine($"Skipping file '{filePath}': {ex.Message}");
                }
                catch (Exception ex)
                {
                    // Catch any other unexpected errors
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            try
            {
                string folderPath = @"C:\ExcelFiles";
                string disclaimerText = "Confidential: This document is for internal use only.";
                string outputFolder = @"C:\ProcessedExcelFiles";

                DisclaimerAdder.ProcessFolder(folderPath, disclaimerText, outputFolder);

                Console.WriteLine("Processing completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fatal error: {ex.Message}");
            }
        }
    }
}
