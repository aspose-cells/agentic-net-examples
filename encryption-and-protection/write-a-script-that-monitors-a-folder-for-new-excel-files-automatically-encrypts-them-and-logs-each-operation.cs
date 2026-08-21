// Title: C# FileSystemWatcher to Auto‑Encrypt New Excel Files with Aspose.Cells and Log Actions
// Description: A C# console app that watches a directory for newly created Excel workbooks, encrypts each file with a strong password using Aspose.Cells, saves the protected copy to a target folder, and writes a timestamped entry to an audit log. It skips non‑Excel files and already‑encrypted workbooks, handling errors gracefully.
// Keywords: Aspose.Cells | C# | FileSystemWatcher | Excel encryption | auto encrypt Excel files | password‑protected workbook | monitor folder for Excel | audit log | detect encrypted workbook | save encrypted copy
// Common Searches: C# monitor folder and encrypt Excel files | Aspose.Cells encrypt workbook automatically | FileSystemWatcher encrypt new Excel files | log Excel encryption events C# | skip already encrypted Excel workbook Aspose
// Developer Intent: Automatically protect every new Excel workbook placed in a watched directory by encrypting it with a password and recording the operation in a log file.
// Use Cases: Secure confidential reports uploaded to a shared drop‑box before distribution. | Encrypt incoming data extracts and store them in an archive for regulatory compliance. | Create an audit trail of file‑level encryption for internal or external audits.
// AI Prompts: Write a C# method that takes a file path and a password, uses Aspose.Cells to encrypt the workbook, and returns the path of the encrypted file. | Provide a resilient FileSystemWatcher wrapper that retries processing a file until it is no longer locked. | Generate a unit test that confirms the script ignores files that are already encrypted.

using System;
using System.IO;
using System.Threading;
using Aspose.Cells;

namespace ExcelFolderEncryptor
{
    // A C# console app that watches a directory for newly created Excel workbooks, encrypts each file with a strong password using Aspose.Cells, saves the protected copy to a target folder, and writes a timestamped entry to an audit log. It skips non‑Excel files and already‑encrypted workbooks, handling errors gracefully.
    class Program
    {
        // Password used for encrypting the workbooks
        private const string EncryptionPassword = "StrongPassword123!";

        // Folder to monitor
        private static readonly string WatchFolder = @"C:\WatchedFolder";

        // Folder where encrypted files will be saved (can be same as WatchFolder)
        private static readonly string OutputFolder = @"C:\EncryptedFolder";

        // Log file path
        private static readonly string LogFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "encryption_log.txt");

        static void Main()
        {
            try
            {
                // Ensure output folder exists
                Directory.CreateDirectory(OutputFolder);
            }
            catch (Exception ex)
            {
                Log($"Failed to create output folder '{OutputFolder}': {ex.Message}");
                return;
            }

            // Set up the FileSystemWatcher
            using (FileSystemWatcher watcher = new FileSystemWatcher())
            {
                watcher.Path = WatchFolder;
                watcher.Filter = "*.*"; // Watch all files, we'll filter by extension later
                watcher.Created += OnCreated;
                watcher.EnableRaisingEvents = true;

                Console.WriteLine($"Monitoring folder: {WatchFolder}");
                Console.WriteLine("Press 'q' to quit.");

                // Keep the application running until user quits
                while (Console.Read() != 'q') ;
            }
        }

        private static void OnCreated(object sender, FileSystemEventArgs e)
        {
            // Process only Excel files
            string extension = Path.GetExtension(e.FullPath).ToLowerInvariant();
            if (extension != ".xlsx" && extension != ".xls" && extension != ".xlsb" && extension != ".xlsm")
            {
                return;
            }

            // Ensure the file actually exists before proceeding
            if (!File.Exists(e.FullPath))
            {
                Log($"File '{e.Name}' does not exist. Skipping.");
                return;
            }

            // Wait briefly to ensure the file is fully written
            Thread.Sleep(500);

            try
            {
                // Detect if the file is already encrypted
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(e.FullPath);
                if (formatInfo.IsEncrypted)
                {
                    Log($"File '{e.Name}' is already encrypted. Skipping.");
                    return;
                }
            }
            catch (Exception ex)
            {
                Log($"Failed to detect format for '{e.Name}': {ex.Message}");
                return;
            }

            try
            {
                // Load the workbook (no password needed)
                if (!File.Exists(e.FullPath))
                {
                    Log($"File '{e.Name}' disappeared before processing.");
                    return;
                }

                Workbook workbook = new Workbook(e.FullPath);

                // Set the password for opening the workbook (encryption)
                workbook.Settings.Password = EncryptionPassword;

                // Set stronger encryption options (ignored for .xlsx/.xlsm but kept for completeness)
                workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

                // Determine output path (overwrite original or place in output folder)
                string outputPath = Path.Combine(OutputFolder, e.Name);

                // Save the encrypted workbook (same format as original)
                SaveFormat saveFormat = GetSaveFormatFromExtension(extension);
                workbook.Save(outputPath, saveFormat);

                Log($"Encrypted and saved file: '{e.Name}' to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Log($"Error processing file '{e.Name}': {ex.Message}");
            }
        }

        // Helper to map file extension to Aspose.Cells SaveFormat
        private static SaveFormat GetSaveFormatFromExtension(string ext)
        {
            return ext switch
            {
                ".xlsx" => SaveFormat.Xlsx,
                ".xlsm" => SaveFormat.Xlsx,
                ".xlsb" => SaveFormat.Xlsb,
                ".xls" => SaveFormat.Excel97To2003,
                _ => SaveFormat.Xlsx,
            };
        }

        // Simple logging to a text file with timestamp
        private static void Log(string message)
        {
            string entry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}";
            Console.WriteLine(entry);
            try
            {
                File.AppendAllText(LogFilePath, entry + Environment.NewLine);
            }
            catch
            {
                // Silently ignore logging failures to keep the watcher alive
            }
        }
    }
}
