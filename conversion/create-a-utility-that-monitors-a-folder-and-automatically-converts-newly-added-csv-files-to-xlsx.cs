// Title: C# program that watches a folder and automatically converts incoming CSV files to XLSX using Aspose.Cells
// Common Searches: C# FileSystemWatcher convert new CSV files to Excel with Aspose.Cells | automatically transform CSV to XLSX on file creation in .NET | watch directory for CSV and save as XLSX using Aspose.Cells library | background thread CSV to Excel conversion example in C# | real‑time folder monitoring and CSV to Excel export in C#
// Tags: FileSystemWatcher CSV to XLSX conversion | Aspose.Cells load CSV workbook | background thread Excel export from CSV | real-time folder monitoring for Excel conversion | automatic CSV to Excel conversion C#

using System;
using System.IO;
using System.Threading;
using Aspose.Cells;

// The utility uses FileSystemWatcher to monitor a specified directory for newly created *.csv files. When a CSV appears, it queues a background task that waits briefly, loads the file into an Aspose.Cells Workbook with CSV load options, and saves it as an .xlsx workbook, logging success or errors to the console.
class CsvToXlsxConverter
{
    // Folder to monitor
    private static readonly string WatchFolder = @"C:\WatchFolder";

    // Time to wait for the file to be fully written (in milliseconds)
    private const int FileReadyDelay = 500;

    static void Main()
    {
        // Ensure the folder exists
        if (!Directory.Exists(WatchFolder))
        {
            Console.WriteLine($"Folder does not exist: {WatchFolder}");
            return;
        }

        // Set up the FileSystemWatcher
        using (FileSystemWatcher watcher = new FileSystemWatcher())
        {
            watcher.Path = WatchFolder;
            watcher.Filter = "*.csv";
            watcher.NotifyFilter = NotifyFilters.FileName | NotifyFilters.CreationTime;

            // Subscribe to the Created event
            watcher.Created += OnCreated;

            // Begin watching
            watcher.EnableRaisingEvents = true;

            Console.WriteLine($"Monitoring folder: {WatchFolder}");
            Console.WriteLine("Press 'q' to quit.");

            // Keep the application running until user quits
            while (Console.Read() != 'q')
            {
                Thread.Sleep(100);
            }
        }
    }

    private static void OnCreated(object sender, FileSystemEventArgs e)
    {
        // Run conversion on a separate thread to avoid blocking the watcher
        ThreadPool.QueueUserWorkItem(state => ConvertCsvToXlsx(e.FullPath));
    }

    private static void ConvertCsvToXlsx(string csvPath)
    {
        try
        {
            // Wait briefly to ensure the file is no longer being written
            Thread.Sleep(FileReadyDelay);

            // Verify the file still exists
            if (!File.Exists(csvPath))
                return;

            // Load the CSV file into an Aspose.Cells Workbook
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook workbook = new Workbook(csvPath, loadOptions);

            // Determine the output XLSX file path
            string xlsxPath = Path.ChangeExtension(csvPath, ".xlsx");

            // Save the workbook as XLSX
            workbook.Save(xlsxPath, SaveFormat.Xlsx);

            Console.WriteLine($"Converted: {Path.GetFileName(csvPath)} -> {Path.GetFileName(xlsxPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error converting file '{csvPath}': {ex.Message}");
        }
    }
}
