// Title: C# .NET scheduled task to monitor a folder and update Excel worksheet TabId with Aspose.Cells
// Description: A console‑based .NET app that uses a System.Threading.Timer to scan a specified directory every few minutes, loads each new *.xlsx, *.xls, *.xlsb, or *.xlsm workbook with Aspose.Cells, assigns a unique Worksheet.TabId (e.g., 1000 + sheet index), saves the file, and tracks processed files to avoid repeats.
// Keywords: Aspose.Cells TabId update | C# folder monitor Excel | scheduled task Excel processing | periodic worksheet TabId assignment | Windows service Aspose.Cells | batch update Excel metadata | file system scanner C# | Excel TabId programmatic set | auto‑process Excel files .NET
// Common Searches: how to automatically set worksheet TabId in C# | C# timer to scan folder for new Excel files | Aspose.Cells update TabId for multiple workbooks | create a Windows service that modifies Excel worksheets | periodic Excel file processing with Aspose.Cells
// Developer Intent: Create a reliable, repeatable task that watches a directory, loads any newly added Excel workbooks, programmatically assigns new TabId values to each worksheet, and saves the changes without re‑processing the same file.
// Use Cases: Automatically assign sequential TabId values to worksheets as soon as reports are dropped into a shared folder. | Run the scanner as a Windows service to keep worksheet identifiers consistent across all Excel files in a data‑pipeline environment. | Integrate the folder‑scan routine into an ETL workflow so that TabId metadata is refreshed before downstream analytics.
// AI Prompts: Generate C# code that uses Aspose.Cells to set Worksheet.TabId based on a custom rule while handling .xls, .xlsx, .xlsb, and .xlsm files. | Provide a robust file‑system watcher implementation that triggers the TabId update only for newly created Excel files, including logging and error handling. | Rewrite the timer‑based scanning loop as a .NET Core hosted background service with dependency injection and graceful shutdown.

using System;
using System.IO;
using System.Collections.Generic;
using System.Threading;
using Aspose.Cells;

namespace WorksheetTabIdUpdater
{
    // A console‑based .NET app that uses a System.Threading.Timer to scan a specified directory every few minutes, loads each new *.xlsx, *.xls, *.xlsb, or *.xlsm workbook with Aspose.Cells, assigns a unique Worksheet.TabId (e.g., 1000 + sheet index), saves the file, and tracks processed files to avoid repeats.
    class Program
    {
        // Folder to monitor
        private static readonly string folderPath = @"C:\ExcelFolder";

        // Interval for scanning (e.g., every 5 minutes)
        private static readonly TimeSpan scanInterval = TimeSpan.FromMinutes(5);

        // Keep track of files already processed in this session
        private static readonly HashSet<string> processedFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        static void Main(string[] args)
        {
            // Initial scan
            ScanAndUpdate();

            // Set up a timer to run the scan periodically
            Timer timer = new Timer(_ => ScanAndUpdate(), null, scanInterval, scanInterval);

            // Prevent the application from exiting
            Console.WriteLine("Press Enter to stop the service...");
            Console.ReadLine();
        }

        private static void ScanAndUpdate()
        {
            try
            {
                // Get all Excel files in the folder (including subfolders if needed)
                string[] excelFiles = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly);
                foreach (string filePath in excelFiles)
                {
                    // Filter by known Excel extensions
                    string ext = Path.GetExtension(filePath).ToLowerInvariant();
                    if (ext != ".xlsx" && ext != ".xls" && ext != ".xlsb" && ext != ".xlsm")
                        continue;

                    // Skip files already processed in this run
                    if (processedFiles.Contains(filePath))
                        continue;

                    // Load the workbook (uses the provided load rule)
                    Workbook workbook = new Workbook(filePath);

                    // Update TabId for each worksheet
                    for (int i = 0; i < workbook.Worksheets.Count; i++)
                    {
                        Worksheet sheet = workbook.Worksheets[i];
                        // Example logic: set TabId to a unique value based on sheet index
                        sheet.TabId = 1000 + i;
                    }

                    // Save the workbook back to the same file (uses the provided save rule)
                    workbook.Save(filePath);

                    // Mark as processed
                    processedFiles.Add(filePath);

                    Console.WriteLine($"Updated TabId for workbook: {Path.GetFileName(filePath)}");
                }
            }
            catch (Exception ex)
            {
                // Log or handle exceptions as needed
                Console.WriteLine($"Error during scan: {ex.Message}");
            }
        }
    }
}
