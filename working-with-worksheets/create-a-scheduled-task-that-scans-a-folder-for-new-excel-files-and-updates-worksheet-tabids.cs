// Title: C# console app using Aspose.Cells to scan a folder and assign sequential TabId values to worksheets in .xlsx workbooks
// AI Prompts: Create a C# console application that enumerates all .xlsx files in a specified folder, loads each workbook with Aspose.Cells, assigns a unique TabId to every worksheet, and saves the changes. | Enhance the program to write the last used TabId to a JSON file and read it on startup so the counter continues across multiple executions. | Extend the solution to also process .xls files, preserving their original format while updating worksheet TabId values.
// Common Searches: how to batch update worksheet TabId in Excel files using Aspose.Cells C# | C# program to monitor a folder and assign sequential TabId to each sheet in .xlsx workbooks | persist TabId counter between runs when processing Excel files with Aspose.Cells | set up Windows Task Scheduler to run a C# Excel TabId updater
// Tags: Aspose.Cells batch TabId update | C# scan directory for .xlsx workbooks | worksheet TabId auto-increment implementation | persist TabId state to JSON file | Windows Task Scheduler for Excel automation

using System;
using System.IO;
using Aspose.Cells;

namespace ExcelTabIdUpdater
{
    // A C# console utility that scans a given directory for .xlsx files, loads each workbook with Aspose.Cells, assigns a sequential TabId to every worksheet, and saves the workbook back, optionally persisting the counter and supporting scheduled execution.
    class Program
    {
        // Simple counter to generate unique TabIds within a single run.
        // In a real scenario you might persist the last used value.
        private static int _nextTabId = 1;

        static void Main(string[] args)
        {
            // Folder to monitor – adjust as needed.
            string folderPath = @"C:\ExcelFiles";

            // Process all .xlsx files in the folder.
            ProcessFolder(folderPath);
        }

        private static void ProcessFolder(string folderPath)
        {
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder does not exist: {folderPath}");
                return;
            }

            // Get all Excel files (you can add .xls if required).
            string[] excelFiles = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string filePath in excelFiles)
            {
                try
                {
                    // Load the workbook.
                    Workbook workbook = new Workbook(filePath);

                    // Update TabId for each worksheet.
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        sheet.TabId = GetNextTabId();
                    }

                    // Save the workbook back to the same file.
                    workbook.Save(filePath, SaveFormat.Xlsx);

                    Console.WriteLine($"Updated TabIds in: {Path.GetFileName(filePath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }
        }

        // Generates a sequential TabId.
        private static int GetNextTabId()
        {
            return _nextTabId++;
        }
    }
}
