using System;
using System.IO;
using Aspose.Cells;

namespace ExcelEncryptor
{
    class Program
    {
        // Password used for encryption – change as needed
        private const string EncryptionPassword = "StrongPassword123";

        static void Main(string[] args)
        {
            // Validate command‑line arguments
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: ExcelEncryptor <directoryPath>");
                return;
            }

            string directoryPath = args[0];

            if (!Directory.Exists(directoryPath))
            {
                Console.WriteLine($"Error: Directory \"{directoryPath}\" does not exist.");
                return;
            }

            // Supported Excel extensions
            string[] extensions = new[] { ".xls", ".xlsx", ".xlsb", ".xlsm", ".ods", ".csv" };

            int totalFiles = 0;
            int encryptedCount = 0;
            int skippedCount = 0;

            foreach (string filePath in Directory.EnumerateFiles(directoryPath, "*.*", SearchOption.AllDirectories))
            {
                if (Array.IndexOf(extensions, Path.GetExtension(filePath).ToLowerInvariant()) < 0)
                    continue; // Not an Excel file

                totalFiles++;

                try
                {
                    // Detect if the file is already encrypted
                    FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
                    if (formatInfo.IsEncrypted)
                    {
                        Console.WriteLine($"[Skipped] Already encrypted: {filePath}");
                        skippedCount++;
                        continue;
                    }

                    // Load the workbook (no password needed because it's not encrypted)
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
                    Workbook workbook = new Workbook(filePath, loadOptions);

                    // Apply encryption settings
                    workbook.Settings.Password = EncryptionPassword;
                    workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

                    // Overwrite the original file with the encrypted version
                    workbook.Save(filePath);

                    Console.WriteLine($"[Encrypted] {filePath}");
                    encryptedCount++;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[Error] Failed to process \"{filePath}\": {ex.Message}");
                }
            }

            // Summary report
            Console.WriteLine();
            Console.WriteLine("=== Encryption Summary ===");
            Console.WriteLine($"Total Excel files found : {totalFiles}");
            Console.WriteLine($"Successfully encrypted   : {encryptedCount}");
            Console.WriteLine($"Skipped (already encrypted) : {skippedCount}");
            Console.WriteLine($"Processing completed at {DateTime.Now}");
        }
    }
}