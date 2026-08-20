// Title: C# Scheduled Folder Monitor that Updates Worksheet TabId with Aspose.Cells
// Description: A .NET console app that uses System.Timers.Timer to scan a specified directory every 5 minutes, loads each .xlsx file with Aspose.Cells, assigns a sequential TabId to every worksheet, saves the changes, and logs the operation. Ideal for automating TabId normalization in shared Excel repositories or background services.
// Keywords: Aspose.Cells TabId update | C# folder monitor Excel | timer based Excel processing | scan directory for .xlsx files | automate worksheet TabId | Windows service Excel automation | C# scheduled task Aspose.Cells
// Common Searches: C# code to monitor a folder and modify Excel files with Aspose.Cells | How to set worksheet TabId programmatically in .NET | Scheduled task for updating Excel TabId values | Aspose.Cells timer example for batch processing | Automatic TabId assignment for Excel workbooks
// Developer Intent: Create an automated process that periodically scans a folder, loads each new Excel workbook, and sets a consistent TabId for every worksheet using Aspose.Cells.
// Use Cases: Ensure daily report workbooks in a shared network folder have sequential TabIds for predictable tab ordering. | Run as a background Windows service that normalizes TabId settings after bulk edits or data imports. | Integrate into a CI/CD pipeline where generated Excel files must have their worksheet TabIds standardized before distribution.
// AI Prompts: Generate C# code that monitors a directory and updates worksheet TabId values every 5 minutes using Aspose.Cells, with robust error handling. | Explain how to modify the timer interval and customize the TabId assignment logic (e.g., based on worksheet name) in the provided example. | Provide recommendations for converting this console app into a Windows Service, including logging, configuration files, and graceful shutdown.

using System;
using System.IO;
using System.Timers;
using Aspose.Cells;

namespace ExcelTabIdUpdater
{
    // A .NET console app that uses System.Timers.Timer to scan a specified directory every 5 minutes, loads each .xlsx file with Aspose.Cells, assigns a sequential TabId to every worksheet, saves the changes, and logs the operation. Ideal for automating TabId normalization in shared Excel repositories or background services.
    class Program
    {
        // Folder to monitor for Excel files
        private static readonly string FolderPath = @"C:\ExcelFolder";

        // Interval for scanning the folder (e.g., every 5 minutes)
        private static readonly double ScanIntervalMs = TimeSpan.FromMinutes(5).TotalMilliseconds;

        static void Main()
        {
            // Ensure the folder exists
            if (!Directory.Exists(FolderPath))
            {
                Console.WriteLine($"Folder does not exist: {FolderPath}");
                return;
            }

            // Set up a timer to run the scan periodically
            System.Timers.Timer timer = new System.Timers.Timer(ScanIntervalMs);
            timer.Elapsed += OnTimerElapsed;
            timer.AutoReset = true;

            try
            {
                timer.Start();
                Console.WriteLine($"Monitoring folder: {FolderPath}");
                Console.WriteLine("Press Enter to stop.");
                Console.ReadLine();
            }
            finally
            {
                timer.Stop();
                timer.Dispose();
            }
        }

        // Timer callback to process all Excel files in the folder
        private static void OnTimerElapsed(object sender, ElapsedEventArgs e)
        {
            try
            {
                ProcessFolder(FolderPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing folder '{FolderPath}': {ex.Message}");
            }
        }

        // Scans the specified folder for .xlsx files and updates each worksheet's TabId
        private static void ProcessFolder(string path)
        {
            string[] excelFiles = Directory.GetFiles(path, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string filePath in excelFiles)
            {
                // Ensure the file still exists before attempting to load
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found (skipped): {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Update TabId for each worksheet (example: set to worksheet index + 1)
                    for (int i = 0; i < workbook.Worksheets.Count; i++)
                    {
                        Worksheet sheet = workbook.Worksheets[i];
                        sheet.TabId = i + 1;
                    }

                    // Save the workbook back to the same file
                    workbook.Save(filePath);
                    Console.WriteLine($"Processed file: {filePath}");
                }
                catch (Exception ex)
                {
                    // Log or handle errors as needed
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }
        }
    }
}
