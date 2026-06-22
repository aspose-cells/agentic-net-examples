using System;
using System.IO;
using System.Threading;
using Aspose.Cells;

class Program
{
    // Path to monitor
    private static readonly string WatchFolder = @"C:\WatchedFolder";

    // Log file inside the watched folder
    private static readonly string LogFile = Path.Combine(WatchFolder, "encryption_log.txt");

    static void Main()
    {
        // Ensure the folder exists
        Directory.CreateDirectory(WatchFolder);

        // Set up a watcher for any new file
        var watcher = new FileSystemWatcher(WatchFolder)
        {
            Filter = "*.*",
            EnableRaisingEvents = true,
            IncludeSubdirectories = false
        };

        // React to created files
        watcher.Created += (s, e) => ProcessNewFile(e.FullPath);

        Console.WriteLine($"Monitoring folder: {WatchFolder}");
        Console.WriteLine("Press ENTER to stop.");
        Console.ReadLine();
    }

    private static void ProcessNewFile(string filePath)
    {
        // Simple retry to ensure the file is fully written
        const int maxAttempts = 5;
        const int delayMs = 500;
        for (int i = 0; i < maxAttempts; i++)
        {
            if (IsFileReady(filePath))
                break;
            Thread.Sleep(delayMs);
        }

        // Only handle Excel workbook extensions
        string ext = Path.GetExtension(filePath).ToLowerInvariant();
        if (ext != ".xlsx" && ext != ".xls" && ext != ".xlsm" && ext != ".xlsb")
            return;

        try
        {
            // Load the workbook (uses the provided load rule)
            var loadOptions = new LoadOptions(); // default load options
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Set a password for opening the workbook
            workbook.Settings.Password = "MySecretPassword";

            // Apply strong encryption (uses the provided SetEncryptionOptions rule)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the encrypted workbook, overwriting the original (uses the provided save rule)
            workbook.Save(filePath);

            // Log successful encryption
            Log($"[{DateTime.UtcNow:u}] Encrypted: {Path.GetFileName(filePath)}");
        }
        catch (Exception ex)
        {
            // Log any error that occurs during processing
            Log($"[{DateTime.UtcNow:u}] Error encrypting {Path.GetFileName(filePath)}: {ex.Message}");
        }
    }

    // Helper to check if a file is ready for reading
    private static bool IsFileReady(string path)
    {
        try
        {
            using (FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.None))
            {
                return stream.Length > 0;
            }
        }
        catch
        {
            return false;
        }
    }

    // Append a line to the log file
    private static void Log(string message)
    {
        File.AppendAllText(LogFile, message + Environment.NewLine);
    }
}