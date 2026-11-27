// Title: Batch update worksheet TabId values in all Excel files of a folder and generate a CSV change log with Aspose.Cells for .NET
// Common Searches: C# script to assign sequential TabId to every sheet in multiple .xlsx files using Aspose.Cells | How to create a summary CSV of worksheet TabId changes after batch processing Excel workbooks in .NET | Automate updating sheet TabId across a directory of Excel workbooks with Aspose.Cells | Aspose.Cells example for iterating worksheets and modifying the TabId property in bulk | Generate a UTF‑8 CSV report of old and new TabId values for Excel sheets processed in C#
// Tags: batch worksheet TabId update Aspose.Cells | CSV change log generation Aspose.Cells | process all Excel files in a directory C# | sequential TabId assignment workbook | overwrite workbook after TabId modification

using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using Aspose.Cells;

namespace TabIdUpdater
{
    // Scans a specified (or current) folder for Excel files (.xls, .xlsx, .xlsb, .xlsm), iterates each worksheet, assigns a globally incremented TabId, saves the workbooks in place, and writes a UTF‑8 CSV summarizing file path, sheet name, old TabId, and new TabId for every change.
    class Program
    {
        // Structure to hold change information for CSV output
        private class TabIdChange
        {
            public string FilePath { get; set; }
            public string SheetName { get; set; }
            public int OldTabId { get; set; }
            public int NewTabId { get; set; }
        }

        static void Main(string[] args)
        {
            // Determine the directory to process
            string targetDirectory;
            if (args.Length > 0 && Directory.Exists(args[0]))
            {
                targetDirectory = args[0];
            }
            else
            {
                targetDirectory = Directory.GetCurrentDirectory();
                Console.WriteLine($"No valid directory argument supplied. Using current directory: {targetDirectory}");
            }

            // Supported Excel extensions
            string[] extensions = new[] { ".xls", ".xlsx", ".xlsb", ".xlsm" };

            // Collect all Excel files recursively
            List<string> excelFiles = new List<string>();
            foreach (var ext in extensions)
            {
                excelFiles.AddRange(Directory.GetFiles(targetDirectory, "*" + ext, SearchOption.AllDirectories));
            }

            if (excelFiles.Count == 0)
            {
                Console.WriteLine("No Excel files found in the specified directory.");
                return;
            }

            // List to store all changes for CSV generation
            List<TabIdChange> changes = new List<TabIdChange>();

            // Global counter for new TabIds (starts at 1)
            int globalTabIdCounter = 1;

            foreach (var filePath in excelFiles)
            {
                Console.WriteLine($"Processing file: {filePath}");
                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Iterate through each worksheet
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        int oldTabId = sheet.TabId;
                        int newTabId = globalTabIdCounter++;

                        // Update the TabId
                        sheet.TabId = newTabId;

                        // Record the change
                        changes.Add(new TabIdChange
                        {
                            FilePath = filePath,
                            SheetName = sheet.Name,
                            OldTabId = oldTabId,
                            NewTabId = newTabId
                        });

                        Console.WriteLine($"  Sheet \"{sheet.Name}\": TabId changed from {oldTabId} to {newTabId}");
                    }

                    // Save the workbook back to the same file (overwrite)
                    workbook.Save(filePath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file \"{filePath}\": {ex.Message}");
                    // Continue with next file
                }
            }

            // Generate summary CSV
            string csvPath = Path.Combine(targetDirectory, "TabIdChangesSummary.csv");
            try
            {
                using (var writer = new StreamWriter(csvPath, false, Encoding.UTF8))
                {
                    // Write header
                    writer.WriteLine("FilePath,SheetName,OldTabId,NewTabId");

                    // Write each change
                    foreach (var change in changes)
                    {
                        // Escape double quotes in file path or sheet name
                        string escapedFilePath = $"\"{change.FilePath.Replace("\"", "\"\"")}\"";
                        string escapedSheetName = $"\"{change.SheetName.Replace("\"", "\"\"")}\"";

                        writer.WriteLine($"{escapedFilePath},{escapedSheetName},{change.OldTabId},{change.NewTabId}");
                    }
                }

                Console.WriteLine($"Summary CSV generated at: {csvPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write summary CSV: {ex.Message}");
            }
        }
    }
}
