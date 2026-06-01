using System;
using System.IO;
using Aspose.Cells.Utility;

namespace CsvToXlsxWatcher
{
    // Utility that watches a folder and converts new CSV files to XLSX using Aspose.Cells ConversionUtility.
    class Program
    {
        // Folder to monitor – change as needed.
        private const string WatchFolder = @"C:\WatchedFolder";

        static void Main()
        {
            // Ensure the folder exists.
            if (!Directory.Exists(WatchFolder))
            {
                Console.WriteLine($"Folder does not exist: {WatchFolder}");
                return;
            }

            // Set up the FileSystemWatcher.
            using (FileSystemWatcher watcher = new FileSystemWatcher())
            {
                watcher.Path = WatchFolder;
                watcher.Filter = "*.csv";               // Watch only CSV files.
                watcher.Created += OnCreated;           // Event raised when a new file appears.
                watcher.EnableRaisingEvents = true;     // Start monitoring.

                Console.WriteLine($"Monitoring folder: {WatchFolder}");
                Console.WriteLine("Press ENTER to exit.");
                Console.ReadLine(); // Keep the application running.
            }
        }

        // Event handler called when a new CSV file is created.
        private static void OnCreated(object sender, FileSystemEventArgs e)
        {
            // Give the OS a moment to release the file lock.
            System.Threading.Thread.Sleep(500);

            string sourcePath = e.FullPath;
            string destPath = Path.ChangeExtension(sourcePath, ".xlsx");

            try
            {
                // Use Aspose.Cells ConversionUtility to convert CSV to XLSX.
                // This follows the provided rule: ConversionUtility.Convert(string, string)
                ConversionUtility.Convert(sourcePath, destPath);
                Console.WriteLine($"Converted: {Path.GetFileName(sourcePath)} -> {Path.GetFileName(destPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error converting '{sourcePath}': {ex.Message}");
            }
        }
    }
}