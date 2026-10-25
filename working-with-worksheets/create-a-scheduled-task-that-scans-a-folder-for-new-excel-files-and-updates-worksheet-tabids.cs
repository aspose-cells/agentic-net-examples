// Title: C# program using Aspose.Cells to monitor a folder and assign sequential TabId values to worksheets in new .xlsx files
// Common Searches: how to monitor a directory for new .xlsx files and update worksheet TabId with Aspose.Cells in C# | C# timer to periodically scan a folder and set sequential TabId for each worksheet using Aspose.Cells | automate worksheet TabId assignment for Excel workbooks added to a folder with Aspose.Cells | using Aspose.Cells to assign incremental TabId to worksheets during batch processing in C# | C# code example for folder watch and worksheet TabId update with Aspose.Cells
// Tags: folder monitoring with Aspose.Cells C# | sequential worksheet TabId assignment Aspose.Cells | periodic Excel file processing timer C# | batch update worksheet TabId Aspose.Cells | in‑memory processed file tracking C#

using System;
using System.IO;
using System.Threading;
using Aspose.Cells;

namespace ExcelTabIdUpdater
{
    // The example defines a folder path and a 5‑minute interval, then uses System.Threading.Timer to repeatedly scan the directory for *.xlsx files. For each unprocessed workbook, it loads the file with Aspose.Cells, assigns incremental TabId values to every worksheet, saves the workbook back to the same file, and records the file path in a HashSet to avoid re‑processing.
    class Program
    {
        // Folder to monitor
        private static readonly string FolderPath = @"C:\ExcelFiles";

        // Interval for scanning (e.g., every 5 minutes)
        private static readonly TimeSpan ScanInterval = TimeSpan.FromMinutes(5);

        // Simple in‑memory store to remember processed files
        private static readonly HashSet<string> ProcessedFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        static void Main(string[] args)
        {
            // Initial scan
            ScanAndUpdate();

            // Set up a timer to run the scan periodically
            Timer timer = new Timer(_ => ScanAndUpdate(), null, ScanInterval, ScanInterval);

            // Prevent the application from exiting
            Console.WriteLine("Excel TabId updater is running. Press Enter to exit.");
            Console.ReadLine();
        }

        private static void ScanAndUpdate()
        {
            try
            {
                // Get all Excel files in the folder (including subfolders if needed)
                var excelFiles = Directory.GetFiles(FolderPath, "*.xlsx", SearchOption.TopDirectoryOnly);

                foreach (var filePath in excelFiles)
                {
                    // Skip files that have already been processed
                    if (ProcessedFiles.Contains(filePath))
                        continue;

                    // Load the workbook (using the provided load rule)
                    Workbook workbook = new Workbook(filePath);

                    // Update TabId for each worksheet
                    int nextTabId = 1;
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        sheet.TabId = nextTabId++;
                    }

                    // Save the workbook (using the provided save rule)
                    workbook.Save(filePath, SaveFormat.Xlsx);

                    // Mark file as processed
                    ProcessedFiles.Add(filePath);
                }
            }
            catch (Exception ex)
            {
                // Log or handle exceptions as appropriate
                Console.WriteLine($"Error during scan: {ex.Message}");
            }
        }
    }
}
