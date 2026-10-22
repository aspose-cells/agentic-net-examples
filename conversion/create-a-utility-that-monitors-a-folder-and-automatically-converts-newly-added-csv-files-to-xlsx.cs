// Title: C# console app that watches a folder and automatically converts incoming CSV files to XLSX with Aspose.Cells
// Common Searches: c# FileSystemWatcher convert new csv files to xlsx using Aspose.Cells | how to automatically convert csv to excel when a file is added to a directory in .net | aspocells example for monitoring a folder and saving csv as xlsx | c# monitor folder for csv and export to xlsx with retry logic | convert csv to xlsx on file creation using Aspose.Cells LoadOptions
// Tags: FileSystemWatcher CSV folder monitoring | Aspose.Cells CSV to XLSX conversion | LoadOptions CSV loading Aspose.Cells | Save workbook as XLSX programmatically | real-time CSV to Excel conversion C#

using System;
using System.IO;
using Aspose.Cells;

// A console application that uses FileSystemWatcher to monitor a specified directory for newly created *.csv files, waits until each file is fully written, loads the CSV into an Aspose.Cells Workbook with LoadOptions, and saves it as an .xlsx file with the same name.
class CsvToXlsxConverter
{
    // Folder to monitor
    private static readonly string WatchFolder = @"C:\WatchFolder";

    // Entry point
    static void Main()
    {
        // Ensure the folder exists
        if (!Directory.Exists(WatchFolder))
        {
            Console.WriteLine($"Folder '{WatchFolder}' does not exist.");
            return;
        }

        // Set up a FileSystemWatcher to monitor CSV files
        using (FileSystemWatcher watcher = new FileSystemWatcher())
        {
            watcher.Path = WatchFolder;
            watcher.Filter = "*.csv";
            watcher.Created += OnCreated;
            watcher.EnableRaisingEvents = true;

            Console.WriteLine($"Monitoring folder: {WatchFolder}");
            Console.WriteLine("Press 'q' to quit.");

            // Keep the application running until user quits
            while (Console.Read() != 'q') ;
        }
    }

    // Event handler triggered when a new CSV file is created
    private static void OnCreated(object sender, FileSystemEventArgs e)
    {
        // Wait for the file to be fully written
        try
        {
            // Simple retry logic
            for (int i = 0; i < 5; i++)
            {
                try
                {
                    using (FileStream stream = File.Open(e.FullPath, FileMode.Open, FileAccess.Read, FileShare.None))
                    {
                        // If we can open the file exclusively, it's ready
                        break;
                    }
                }
                catch (IOException)
                {
                    System.Threading.Thread.Sleep(500);
                }
            }

            ConvertCsvToXlsx(e.FullPath);
            Console.WriteLine($"Converted: {e.Name}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing file '{e.Name}': {ex.Message}");
        }
    }

    // Performs the actual conversion using Aspose.Cells
    private static void ConvertCsvToXlsx(string csvPath)
    {
        // Load the CSV file into a Workbook
        Workbook workbook = new Workbook(csvPath, new LoadOptions(LoadFormat.CSV));

        // Determine output path (same name with .xlsx extension)
        string xlsxPath = Path.ChangeExtension(csvPath, ".xlsx");

        // Save the workbook as XLSX
        workbook.Save(xlsxPath, SaveFormat.Xlsx);
    }
}
