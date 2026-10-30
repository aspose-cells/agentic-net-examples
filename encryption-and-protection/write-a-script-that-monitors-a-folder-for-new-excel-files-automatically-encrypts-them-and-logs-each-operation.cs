// Title: C# program that monitors a directory, encrypts incoming .xlsx workbooks with Aspose.Cells password protection, and records each action to a log file
// Common Searches: how to automatically encrypt newly created .xlsx files using Aspose.Cells in C# | C# FileSystemWatcher example for applying password protection to Excel workbooks | log each Excel encryption event to a text file in a C# application | watch folder for Excel files and save them with password using Aspose.Cells | encrypt Excel workbook on creation and append operation details to a log
// Tags: folder monitoring encrypt xlsx Aspose.Cells | password protect Excel workbook C# | file system watcher encrypt excel files | operation logging for encrypted workbooks | Aspose.Cells saveformat xlsx with password

using System;
using System.IO;
using System.Threading;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The C# application watches a specified folder for new .xlsx files, loads each workbook with Aspose.Cells, applies password protection via workbook settings, overwrites the original file with the encrypted version, and appends a timestamped entry to a log file for every successful or failed encryption.
class ExcelEncryptor
{
    // Folder to monitor for new Excel files
    static readonly string folderToWatch = @"C:\WatchFolder";

    // Log file path
    static readonly string logFile = Path.Combine(folderToWatch, "encrypt_log.txt");

    // Password used for encryption
    static readonly string password = "Secret123";

    static void Main()
    {
        try
        {
            // Ensure the watch folder exists
            if (!Directory.Exists(folderToWatch))
            {
                Log($"{DateTime.Now}: Watch folder does not exist: {folderToWatch}");
                return;
            }

            // Set up a FileSystemWatcher to monitor the folder
            using (FileSystemWatcher watcher = new FileSystemWatcher())
            {
                watcher.Path = folderToWatch;
                watcher.Filter = "*.xlsx";               // Watch only .xlsx files
                watcher.Created += OnCreated;            // Event when a new file appears
                watcher.EnableRaisingEvents = true;      // Start watching

                Console.WriteLine($"Monitoring folder: {folderToWatch}");
                Console.WriteLine("Press 'q' then Enter to quit.");

                // Keep the application running until user quits
                while (Console.ReadLine() != "q") { }
            }
        }
        catch (Exception ex)
        {
            Log($"{DateTime.Now}: Fatal error in Main. {ex.Message}");
        }
    }

    // Event handler triggered when a new Excel file is created
    private static void OnCreated(object source, FileSystemEventArgs e)
    {
        // Small delay to ensure the file is completely written
        Thread.Sleep(500);

        // Verify that the file actually exists before processing
        if (!File.Exists(e.FullPath))
        {
            Log($"{DateTime.Now}: File not found '{e.Name}'. Skipping.");
            return;
        }

        try
        {
            // Load the workbook using Aspose.Cells
            Workbook workbook = new Workbook(e.FullPath);

            // Apply password protection (encryption) via workbook settings
            workbook.Settings.Password = password;

            // Overwrite the original file with the encrypted version
            workbook.Save(e.FullPath, SaveFormat.Xlsx);

            Log($"{DateTime.Now}: Encrypted file '{e.Name}' successfully.");
        }
        catch (Exception ex)
        {
            Log($"{DateTime.Now}: Failed to encrypt file '{e.Name}'. Error: {ex.Message}");
        }
    }

    // Simple logging method that appends messages to the log file
    private static void Log(string message)
    {
        try
        {
            File.AppendAllText(logFile, message + Environment.NewLine);
        }
        catch
        {
            // Suppress any logging errors to avoid crashing the watcher
        }
    }
}
