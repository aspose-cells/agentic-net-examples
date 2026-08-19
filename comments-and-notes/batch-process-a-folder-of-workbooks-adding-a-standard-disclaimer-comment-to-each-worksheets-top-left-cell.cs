// Title: Batch add a disclaimer comment to every worksheet in a folder of Excel files with Aspose.Cells (C#)
// Description: C# program that scans a directory for .xlsx, .xls, .xlsm and .xlsb workbooks, opens each file with Aspose.Cells, inserts a visible comment containing a custom disclaimer into cell A1 of every worksheet, and saves the changes back to the original files.
// Keywords: Aspose.Cells | C# Excel comment | batch add comment | disclaimer comment Excel | process multiple workbooks | add comment to A1 | iterate worksheets | Excel file batch processing | Aspose.Cells .NET | visible comment
// Common Searches: asp​ose.cells add same comment to all sheets in a folder | c# batch insert disclaimer into Excel workbooks | add visible comment to cell A1 for every worksheet | process all excel files in a directory with Aspose.Cells | automate comment insertion across multiple Excel files
// Developer Intent: Insert an identical visible disclaimer comment into cell A1 of each worksheet across all Excel files located in a given folder using Aspose.Cells for .NET.
// Use Cases: Automatically tag internal reports with a confidentiality notice before distribution. | Apply a legal disclaimer to every template stored in a shared repository to ensure compliance. | Provide end‑user guidance by adding a standard instruction comment to each sheet of batch‑generated workbooks.
// AI Prompts: Generate C# code with Aspose.Cells that adds a custom comment to cell A1 of every worksheet in all Excel files of a folder, supporting .xlsx, .xls, .xlsm, and .xlsb formats. | Enhance the batch disclaimer script with robust error handling, logging of processed and failed files, and skipping of read‑only workbooks. | Show how to create the disclaimer comment as hidden by default and later toggle its visibility via a macro or additional Aspose.Cells code.

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

namespace BatchWorkbookProcessor
{
    // C# program that scans a directory for .xlsx, .xls, .xlsm and .xlsb workbooks, opens each file with Aspose.Cells, inserts a visible comment containing a custom disclaimer into cell A1 of every worksheet, and saves the changes back to the original files.
    public static class DisclaimerAdder
    {
        /// <param name="folderPath">Path to the folder containing Excel files.</param>
        /// <param name="disclaimer">The disclaimer text to insert as a comment.</param>
        public static void AddDisclaimerToFolder(string folderPath, string disclaimer)
        {
            if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder not found: {folderPath}");
                return;
            }

            // Supported Excel extensions
            string[] extensions = new[] { "*.xlsx", "*.xls", "*.xlsm", "*.xlsb" };

            // Gather all matching files from the folder (non‑recursive)
            List<string> files = new List<string>();
            foreach (var ext in extensions)
            {
                files.AddRange(Directory.GetFiles(folderPath, ext, SearchOption.TopDirectoryOnly));
            }

            foreach (var filePath in files)
            {
                try
                {
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found (skipped): {filePath}");
                        continue;
                    }

                    // Load the workbook using the constructor that accepts a file path
                    using (Workbook workbook = new Workbook(filePath))
                    {
                        // Iterate through all worksheets
                        foreach (Worksheet sheet in workbook.Worksheets)
                        {
                            // Add a comment to cell A1 (top‑left cell)
                            int commentIndex = sheet.Comments.Add("A1");
                            Comment comment = sheet.Comments[commentIndex];
                            comment.Note = disclaimer;
                            comment.Author = "System";

                            // Make the comment visible
                            comment.IsVisible = true;
                        }

                        // Save the modified workbook back to the original file
                        workbook.Save(filePath);
                    }

                    Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }
        }

        // Example usage
        public static void Main()
        {
            try
            {
                string folder = @"C:\ExcelFiles"; // Change to your folder path
                string disclaimerText = "Confidential: This document is for internal use only.";

                AddDisclaimerToFolder(folder, disclaimerText);

                Console.WriteLine("Disclaimer added to all workbooks in the folder.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
