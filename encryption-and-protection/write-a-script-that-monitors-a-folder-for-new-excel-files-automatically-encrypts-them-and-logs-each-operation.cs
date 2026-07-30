// Title: C# FileSystemWatcher to Auto‑Encrypt New Excel Files with Aspose.Cells and Log Operations
// Description: A .NET console app that watches a folder for *.xls* files, skips already‑encrypted workbooks, applies AES‑128 password protection using Aspose.Cells, saves the file, and records each action to the console and a log file.
// Keywords: Aspose.Cells | C# | FileSystemWatcher | Excel encryption | AES 128 | auto encrypt workbook | folder monitor | encryption log | detect encrypted Excel | password protection
// Common Searches: C# monitor folder encrypt Excel with Aspose.Cells | auto encrypt new Excel files .NET | FileSystemWatcher encrypt workbook example | log Excel encryption actions in C# | detect if Excel file is encrypted before saving
// Developer Intent: Automatically secure every new Excel file placed in a designated directory and keep an audit trail of the encryption process.
// Use Cases: Secure client‑generated reports the moment they land in a shared drop folder. | Encrypt financial spreadsheets before archiving them to a network share. | Create a compliance‑ready audit log of all encrypted workbooks.
// AI Prompts: Write a C# method that uses Aspose.Cells to encrypt a workbook with a given password and AES‑128 encryption. | Generate a resilient FileSystemWatcher example that retries when a file is locked and writes failures to a log file. | Provide unit tests for the encryption and logging logic using a temporary directory and mock Excel files.

using System;
using System.IO;
using System.Threading;
using Aspose.Cells;

// A .NET console app that watches a folder for *.xls* files, skips already‑encrypted workbooks, applies AES‑128 password protection using Aspose.Cells, saves the file, and records each action to the console and a log file.
class Program
{
    // Folder to monitor
    private static readonly string WatchFolder = @"C:\WatchedFolder";

    // Password used for encryption
    private static readonly string EncryptionPassword = "StrongPassword123!";

    static void Main()
    {
        // Ensure the folder exists
        if (!Directory.Exists(WatchFolder))
        {
            Console.WriteLine($"Creating folder: {WatchFolder}");
            Directory.CreateDirectory(WatchFolder);
        }

        // Set up a FileSystemWatcher to monitor Excel files
        using (FileSystemWatcher watcher = new FileSystemWatcher(WatchFolder, "*.xls*"))
        {
            watcher.Created += OnCreated;
            watcher.EnableRaisingEvents = true;

            Console.WriteLine($"Monitoring folder: {WatchFolder}");
            Console.WriteLine("Press ENTER to exit.");
            Console.ReadLine(); // Keep the application running
        }
    }

    private static void OnCreated(object sender, FileSystemEventArgs e)
    {
        // Give the OS some time to finish writing the file
        Thread.Sleep(500);

        try
        {
            // Verify the file still exists
            if (!File.Exists(e.FullPath))
            {
                Log($"File not found: {e.FullPath}");
                return;
            }

            // Detect if the file is already encrypted
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(e.FullPath);
            if (formatInfo.IsEncrypted)
            {
                Log($"File already encrypted, skipping: {e.Name}");
                return;
            }

            // Load the workbook (using the standard load rule)
            Workbook workbook = new Workbook(e.FullPath);

            // Apply encryption password
            workbook.Settings.Password = EncryptionPassword;

            // Set encryption options (strong AES 128‑bit)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the workbook back to the same path (using the standard save rule)
            workbook.Save(e.FullPath);

            Log($"Encrypted and saved: {e.Name}");
        }
        catch (Exception ex)
        {
            Log($"Error processing file {e.Name}: {ex.Message}");
        }
    }

    private static void Log(string message)
    {
        string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}";
        Console.WriteLine(logEntry);

        // Append to a simple log file in the watch folder
        string logPath = Path.Combine(WatchFolder, "encryption_log.txt");
        try
        {
            File.AppendAllText(logPath, logEntry + Environment.NewLine);
        }
        catch
        {
            // If logging to file fails, we silently ignore to avoid crashing the watcher
        }
    }
}
