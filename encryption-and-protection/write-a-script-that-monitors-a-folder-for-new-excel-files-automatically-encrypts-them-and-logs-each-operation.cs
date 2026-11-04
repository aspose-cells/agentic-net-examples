// Title: C# FileSystemWatcher that encrypts incoming .xlsx files with Aspose.Cells and logs each operation
// Common Searches: C# FileSystemWatcher encrypt new Excel files using Aspose.Cells password protection | How to automatically apply a password to .xlsx files when they are added to a folder in C# | Log encryption actions for Excel files processed by Aspose.Cells in a .NET application | Retry opening newly created Excel file until it is ready before applying Aspose.Cells encryption | Save encrypted workbook back to original path with Aspose.Cells in C#
// Tags: folder monitoring for new .xlsx files | Aspose.Cells password protection for workbooks | automated Excel encryption in .NET | operation logging to a text file | file readiness retry loop in C#

using System;
using System.IO;
using System.Threading;
using Aspose.Cells;

// The example creates a FileSystemWatcher that watches a specified directory for newly created .xlsx files, waits until each file is fully written, loads it with Aspose.Cells, applies password protection via Workbook.Settings, saves the encrypted file back to its original location, and appends a timestamped entry to an operations.log file indicating success or any errors.
class Program
{
    static void Main()
    {
        // Folder to monitor
        string watchPath = @"C:\WatchFolder";

        // Log file path
        string logPath = Path.Combine(watchPath, "operations.log");

        // Ensure the folder exists
        Directory.CreateDirectory(watchPath);

        // Set up a FileSystemWatcher to detect new .xlsx files
        using (FileSystemWatcher watcher = new FileSystemWatcher())
        {
            watcher.Path = watchPath;
            watcher.Filter = "*.xlsx";
            watcher.Created += (sender, e) => OnCreated(e, logPath);
            watcher.EnableRaisingEvents = true;

            Console.WriteLine($"Monitoring folder: {watchPath}");
            Console.WriteLine("Press ENTER to exit.");
            Console.ReadLine(); // Keep the application running
        }
    }

    // Event handler for newly created files
    private static void OnCreated(FileSystemEventArgs e, string logPath)
    {
        // Wait until the file is fully written and can be accessed exclusively
        const int maxAttempts = 10;
        const int delayMs = 500;
        bool fileReady = false;

        for (int i = 0; i < maxAttempts; i++)
        {
            try
            {
                using (FileStream stream = File.Open(e.FullPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    fileReady = true;
                    break;
                }
            }
            catch (IOException)
            {
                Thread.Sleep(delayMs);
            }
        }

        if (!fileReady)
        {
            LogOperation(logPath, $"[ERROR] File not ready for processing: {e.Name} at {DateTime.Now}");
            return;
        }

        try
        {
            // Load the workbook using Aspose.Cells
            Workbook workbook = new Workbook(e.FullPath);

            // Apply password protection (encryption)
            workbook.Settings.Password = "StrongPassword!123";

            // Save the workbook, overwriting the original file
            workbook.Save(e.FullPath, SaveFormat.Xlsx);

            // Log successful encryption
            LogOperation(logPath, $"[INFO] Encrypted file: {e.Name} at {DateTime.Now}");
        }
        catch (Exception ex)
        {
            // Log any errors that occur during encryption
            LogOperation(logPath, $"[ERROR] Failed to encrypt {e.Name}: {ex.Message} at {DateTime.Now}");
        }
    }

    // Simple logging helper
    private static void LogOperation(string logFilePath, string message)
    {
        try
        {
            File.AppendAllText(logFilePath, message + Environment.NewLine);
        }
        catch
        {
            // If logging fails, swallow the exception to avoid crashing the watcher
        }
    }
}
