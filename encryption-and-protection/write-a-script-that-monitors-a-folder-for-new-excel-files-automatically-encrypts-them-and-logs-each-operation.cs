// Title: C# FileSystemWatcher to Auto‑Encrypt New Excel Files with Aspose.Cells and Log Operations
// Description: A .NET console app that watches a directory for incoming *.xls* files, opens each workbook with Aspose.Cells, applies AES‑128 password protection, overwrites the original file, and appends a timestamped entry to a log file. Includes retry logic for file‑access latency and configurable password and folder paths.
// Keywords: Aspose.Cells encrypt workbook C# | FileSystemWatcher monitor folder | auto encrypt Excel files .NET | AES 128 Excel password protection | C# log encryption activity | watch folder for new Excel files | Windows file watcher encryption | batch encrypt Excel workbooks
// Common Searches: how to encrypt Excel files automatically in C# | C# FileSystemWatcher encrypt and log Excel workbooks | Aspose.Cells set workbook password programmatically | monitor folder and apply AES encryption to Excel files | log file for encrypted Excel documents .NET
// Developer Intent: Automatically secure every new Excel file placed in a designated folder and keep a persistent audit log of the encryption process.
// Use Cases: Secure incoming financial statements saved to a shared drop‑box by encrypting them instantly. | Apply password‑based protection to exported data sets before archiving to meet compliance policies. | Create an audit trail of encrypted workbooks for regulatory reporting, including timestamps and file names.
// AI Prompts: Generate C# code that uses Aspose.Cells to apply AES‑128 encryption with a configurable password to a workbook and overwrite the source file. | Provide a robust FileSystemWatcher example that retries opening a newly created file until it is ready, then encrypts it and writes a log entry. | Suggest enhancements to the logging mechanism to capture file size, encryption duration, and full exception stack traces.

using System;
using System.IO;
using System.Threading;
using Aspose.Cells;

namespace ExcelFolderEncryptor
{
    // A .NET console app that watches a directory for incoming *.xls* files, opens each workbook with Aspose.Cells, applies AES‑128 password protection, overwrites the original file, and appends a timestamped entry to a log file. Includes retry logic for file‑access latency and configurable password and folder paths.
    class Program
    {
        // Password used for encryption – change as needed
        private const string EncryptionPassword = "StrongPassword123";

        // Folder to monitor – change to your target directory
        private static readonly string WatchFolder = @"C:\Temp\ExcelWatch";

        // Simple log file
        private static readonly string LogFilePath = Path.Combine(WatchFolder, "encryption_log.txt");

        static void Main()
        {
            // Ensure the watch folder exists
            if (!Directory.Exists(WatchFolder))
            {
                Console.WriteLine($"Creating watch folder: {WatchFolder}");
                Directory.CreateDirectory(WatchFolder);
            }

            // Set up the FileSystemWatcher
            using (FileSystemWatcher watcher = new FileSystemWatcher())
            {
                watcher.Path = WatchFolder;
                watcher.Filter = "*.xls*";               // Watch all Excel file extensions
                watcher.Created += OnCreated;
                watcher.EnableRaisingEvents = true;

                Console.WriteLine($"Monitoring folder: {WatchFolder}");
                Console.WriteLine("Press 'q' to quit.");

                // Keep the application running until user quits
                while (Console.Read() != 'q') ;
            }
        }

        // Event handler for newly created files
        private static void OnCreated(object sender, FileSystemEventArgs e)
        {
            // Run processing on a separate thread to avoid blocking the watcher
            ThreadPool.QueueUserWorkItem(_ => ProcessFile(e.FullPath));
        }

        // Core logic: load, encrypt, save, and log
        private static void ProcessFile(string filePath)
        {
            // Wait until the file is ready for reading (handles copy/write latency)
            const int maxAttempts = 10;
            int attempt = 0;
            while (attempt < maxAttempts)
            {
                try
                {
                    using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.None))
                    {
                        // If we can open the file exclusively, it is ready
                        break;
                    }
                }
                catch (IOException)
                {
                    attempt++;
                    Thread.Sleep(500); // Wait half a second before retrying
                }
            }

            try
            {
                // Load the workbook (using the standard constructor – lifecycle rule)
                Workbook workbook = new Workbook(filePath);

                // Set the password that protects opening the workbook
                workbook.Settings.Password = EncryptionPassword;

                // Apply encryption options (strong AES 128‑bit)
                workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

                // Save back to the same file (overwrite) – lifecycle rule
                workbook.Save(filePath);

                // Log success
                string message = $"{DateTime.Now}: Encrypted file '{Path.GetFileName(filePath)}' successfully.";
                Console.WriteLine(message);
                Log(message);
            }
            catch (Exception ex)
            {
                // Log any errors
                string error = $"{DateTime.Now}: Failed to encrypt '{Path.GetFileName(filePath)}'. Error: {ex.Message}";
                Console.WriteLine(error);
                Log(error);
            }
        }

        // Append a line to the log file
        private static void Log(string text)
        {
            try
            {
                File.AppendAllText(LogFilePath, text + Environment.NewLine);
            }
            catch
            {
                // If logging fails, we silently ignore to avoid crashing the watcher
            }
        }
    }
}
