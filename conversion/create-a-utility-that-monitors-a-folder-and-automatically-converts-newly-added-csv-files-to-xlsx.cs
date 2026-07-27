// Title: C# FileSystemWatcher: Auto‑convert CSV files to XLSX using Aspose.Cells
// Description: A ready‑to‑run example that watches a folder for new *.csv files, waits until each file is fully written, and instantly converts it to an .xlsx workbook with Aspose.Cells.Utility.ConversionUtility. The conversion runs on a thread‑pool thread, logs success or errors, and can be started or stopped from a console application.
// Keywords: Aspose.Cells CSV to XLSX | C# FileSystemWatcher conversion | auto convert CSV to Excel .NET | folder monitoring Aspose.Cells | ConversionUtility example | real‑time CSV to XLSX | C# background service Excel export
// Common Searches: How to watch a folder and convert CSV to XLSX in C# | Aspose.Cells example for automatic CSV to Excel conversion | FileSystemWatcher convert CSV to Excel on creation | C# code to monitor drop folder and generate XLSX files | Convert CSV to XLSX using Aspose.Cells Utility
// Developer Intent: Automatically transform any CSV file dropped into a specified directory into an Excel workbook using Aspose.Cells.
// Use Cases: Live ingestion of CSV data feeds into Excel for downstream reporting. | Zero‑touch processing of exported CSV files from third‑party systems. | Background service that watches a drop folder, converts incoming CSVs to XLSX, and archives the results for analytics.
// AI Prompts: Create a method that moves the generated XLSX file to an "Archive" subfolder after conversion. | Add exponential‑backoff retry logic to handle transient file‑access errors during conversion. | Write unit tests for CsvToXlsxWatcher using temporary files and a mock FileSystemWatcher to verify the conversion trigger.

using System;
using System.IO;
using System.Threading;
using Aspose.Cells.Utility;

namespace AsposeCellsFolderWatcher
{
    // A ready‑to‑run example that watches a folder for new *.csv files, waits until each file is fully written, and instantly converts it to an .xlsx workbook with Aspose.Cells.Utility.ConversionUtility. The conversion runs on a thread‑pool thread, logs success or errors, and can be started or stopped from a console application.
    public class CsvToXlsxWatcher
    {
        private readonly string _watchFolder;
        private readonly FileSystemWatcher _watcher;

        public CsvToXlsxWatcher(string watchFolder)
        {
            _watchFolder = watchFolder;

            // Initialize the FileSystemWatcher to monitor only *.csv files.
            _watcher = new FileSystemWatcher(_watchFolder, "*.csv")
            {
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.CreationTime,
                EnableRaisingEvents = false,
                IncludeSubdirectories = false
            };

            // Subscribe to the Created event.
            _watcher.Created += OnCsvCreated;
        }

        /// <summary>
        /// Starts monitoring the folder.
        /// </summary>
        public void Start()
        {
            Console.WriteLine($"Starting CSV to XLSX watcher on folder: {_watchFolder}");
            _watcher.EnableRaisingEvents = true;
        }

        /// <summary>
        /// Stops monitoring the folder.
        /// </summary>
        public void Stop()
        {
            _watcher.EnableRaisingEvents = false;
            _watcher.Created -= OnCsvCreated;
            _watcher.Dispose();
            Console.WriteLine("Watcher stopped.");
        }

        // Event handler invoked when a new CSV file appears.
        private void OnCsvCreated(object sender, FileSystemEventArgs e)
        {
            // Run conversion on a separate thread to avoid blocking the watcher.
            ThreadPool.QueueUserWorkItem(_ => ConvertCsvToXlsx(e.FullPath));
        }

        // Performs the actual conversion using Aspose.Cells ConversionUtility.
        private void ConvertCsvToXlsx(string csvPath)
        {
            try
            {
                // Ensure the file is accessible (it may still be being written).
                WaitForFileReady(csvPath);

                // Determine the output XLSX file path (same name, .xlsx extension).
                string xlsxPath = Path.ChangeExtension(csvPath, ".xlsx");

                // Use the provided ConversionUtility.Convert method (source, destination).
                ConversionUtility.Convert(csvPath, xlsxPath);

                Console.WriteLine($"Converted '{Path.GetFileName(csvPath)}' to '{Path.GetFileName(xlsxPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error converting '{Path.GetFileName(csvPath)}': {ex.Message}");
            }
        }

        // Helper method that waits until the file can be opened for reading.
        private void WaitForFileReady(string filePath)
        {
            const int maxAttempts = 10;
            const int delayMs = 500;

            for (int i = 0; i < maxAttempts; i++)
            {
                try
                {
                    using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                    {
                        if (stream.Length > 0)
                            return; // File is ready.
                    }
                }
                catch (IOException)
                {
                    // File is still locked; wait and retry.
                }

                Thread.Sleep(delayMs);
            }

            throw new IOException($"File '{filePath}' is not ready for reading after multiple attempts.");
        }
    }

    // Example program demonstrating the watcher.
    class Program
    {
        static void Main(string[] args)
        {
            // Specify the folder to monitor. Adjust as needed.
            string folderToWatch = Path.Combine(Environment.CurrentDirectory, "CsvInput");

            // Ensure the folder exists.
            Directory.CreateDirectory(folderToWatch);

            var watcher = new CsvToXlsxWatcher(folderToWatch);
            watcher.Start();

            Console.WriteLine("Press ENTER to exit.");
            Console.ReadLine();

            watcher.Stop();
        }
    }
}
