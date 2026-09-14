// Title: C# folder monitor that auto‑converts incoming CSV files to XLSX using Aspose.Cells ConversionUtility
// AI Prompts: Implement a C# folder monitor that detects newly created *.csv files, verifies the file is ready, and invokes Aspose.Cells ConversionUtility to generate an .xlsx workbook in the same directory. | Add robust error handling and logging to capture conversion failures and file‑access issues during the automated CSV‑to‑Excel workflow. | Write a console program that starts the watcher, keeps it running until the user presses ENTER, and gracefully stops the watcher on exit.
// Common Searches: c# monitor folder for new csv files and convert to xlsx with aspose.cells | using FileSystemWatcher to automatically transform csv to excel workbook in .net | asp.net core background service convert csv to xlsx when file is added | check if csv file is fully written before processing in c# | aspocells conversionutility example for csv to ooxml conversion
// Tags: c# directory monitor csv to excel conversion | aspocells conversionutility csv loadoptions | automated ooxml generation from csv c# | file readiness check before aspose conversion | console watcher lifecycle aspocells

using System;
using System.IO;
using System.Threading;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace CsvToXlsxWatcherApp
{
    // Utility class that watches a folder and converts new CSV files to XLSX
    // The example defines a CsvToXlsxWatcher class that uses FileSystemWatcher to watch a specified directory for newly created CSV files, waits until each file is fully written, and then converts it to an XLSX workbook with Aspose.Cells ConversionUtility (CSV load options and OOXML save options). A console Program demonstrates starting, stopping, and logging the watcher’s activity.
    public class CsvToXlsxWatcher
    {
        private readonly string _folderPath;
        private readonly FileSystemWatcher _watcher;

        public CsvToXlsxWatcher(string folderPath)
        {
            _folderPath = folderPath;

            // Initialize FileSystemWatcher to monitor *.csv files
            _watcher = new FileSystemWatcher(_folderPath, "*.csv")
            {
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.CreationTime,
                EnableRaisingEvents = false,
                IncludeSubdirectories = false
            };

            // Subscribe to the Created event
            _watcher.Created += OnCsvCreated;
        }

        // Starts monitoring the folder
        public void Start()
        {
            Console.WriteLine($"Starting CSV to XLSX watcher on folder: {_folderPath}");
            _watcher.EnableRaisingEvents = true;
        }

        // Stops monitoring the folder
        public void Stop()
        {
            _watcher.EnableRaisingEvents = false;
            Console.WriteLine("Watcher stopped.");
        }

        // Event handler triggered when a new CSV file appears
        private void OnCsvCreated(object sender, FileSystemEventArgs e)
        {
            // Run conversion on a separate thread to avoid blocking the watcher
            ThreadPool.QueueUserWorkItem(_ => ProcessCsvFile(e.FullPath));
        }

        // Handles the conversion logic using Aspose.Cells ConversionUtility
        private void ProcessCsvFile(string csvFilePath)
        {
            try
            {
                // Wait briefly to ensure the file is fully written
                for (int i = 0; i < 5; i++)
                {
                    if (IsFileReady(csvFilePath))
                        break;
                    Thread.Sleep(500);
                }

                if (!IsFileReady(csvFilePath))
                {
                    Console.WriteLine($"File not ready for conversion: {csvFilePath}");
                    return;
                }

                // Determine output XLSX path (same name, .xlsx extension)
                string xlsxFilePath = Path.ChangeExtension(csvFilePath, ".xlsx");

                // Create load options for CSV format
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);

                // Create save options for XLSX (OOXML) format
                SaveOptions saveOptions = new OoxmlSaveOptions();

                // Perform conversion using the provided rule
                ConversionUtility.Convert(csvFilePath, loadOptions, xlsxFilePath, saveOptions);

                Console.WriteLine($"Converted '{Path.GetFileName(csvFilePath)}' to '{Path.GetFileName(xlsxFilePath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error converting file '{csvFilePath}': {ex.Message}");
            }
        }

        // Checks whether the file can be opened for reading (i.e., not locked)
        private bool IsFileReady(string filePath)
        {
            try
            {
                using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    return stream.Length > 0;
                }
            }
            catch
            {
                return false;
            }
        }
    }

    // Example program entry point
    class Program
    {
        static void Main(string[] args)
        {
            // Specify the folder to monitor (adjust as needed)
            string folderToWatch = @"C:\Temp\CsvWatchFolder";

            // Ensure the folder exists
            if (!Directory.Exists(folderToWatch))
                Directory.CreateDirectory(folderToWatch);

            // Create and start the watcher
            CsvToXlsxWatcher watcher = new CsvToXlsxWatcher(folderToWatch);
            watcher.Start();

            Console.WriteLine("Press ENTER to exit.");
            Console.ReadLine();

            // Clean up
            watcher.Stop();
        }
    }
}
