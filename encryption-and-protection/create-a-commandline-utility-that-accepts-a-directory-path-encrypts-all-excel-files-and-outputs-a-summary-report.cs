// Title: Create a C# console application that recursively encrypts Excel workbooks in a folder using Aspose.Cells and prints an encryption summary
// Common Searches: C# console program to encrypt all .xlsx files in a directory with Aspose.Cells | How to batch protect Excel workbooks with a password using Aspose.Cells in .NET | Recursive folder scan for Excel files and apply Workbook.Settings.Password in C# | Command‑line tool to set a fixed password on .xls, .xlsx, .xlsm, .xlsb files | Aspose.Cells encrypt multiple Excel files and generate a summary report
// Tags: batch Excel encryption Aspose.Cells | Workbook.Settings.Password C# | recursive Excel file processing .NET | command line Excel password protection | encryption summary report console app

using System;
using System.IO;
using Aspose.Cells;

namespace ExcelEncryptor
{
    // The utility accepts a directory path as a command‑line argument, recursively locates all Excel files (.xls, .xlsx, .xlsm, .xlsb), loads each workbook with Aspose.Cells, assigns a fixed password via Workbook.Settings.Password, saves the workbook overwriting the original, and tracks total, successful, and failed encryptions. At the end it prints a concise summary report to the console.
    class Program
    {
        // Fixed password for encryption – modify as needed
        private const string EncryptionPassword = "Secret123";

        static void Main(string[] args)
        {
            // Validate input arguments
            if (args.Length != 1)
            {
                Console.WriteLine("Usage: ExcelEncryptor <directory_path>");
                return;
            }

            string directoryPath = args[0];

            if (!Directory.Exists(directoryPath))
            {
                Console.WriteLine($"Error: Directory \"{directoryPath}\" does not exist.");
                return;
            }

            // Supported Excel extensions
            string[] extensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb" };

            // Summary counters
            int totalFiles = 0;
            int encryptedCount = 0;
            int failedCount = 0;

            // Iterate through all files with supported extensions
            foreach (string filePath in Directory.EnumerateFiles(directoryPath, "*.*", SearchOption.AllDirectories))
            {
                if (Array.Exists(extensions, ext => ext.Equals(Path.GetExtension(filePath), StringComparison.OrdinalIgnoreCase)))
                {
                    totalFiles++;
                    try
                    {
                        // Load the workbook
                        Workbook workbook = new Workbook(filePath);

                        // Set the password to encrypt the file
                        workbook.Settings.Password = EncryptionPassword;

                        // Save the workbook, overwriting the original file
                        workbook.Save(filePath);

                        encryptedCount++;
                        Console.WriteLine($"Encrypted: {filePath}");
                    }
                    catch (Exception ex)
                    {
                        failedCount++;
                        Console.WriteLine($"Failed to encrypt \"{filePath}\": {ex.Message}");
                    }
                }
            }

            // Output summary report
            Console.WriteLine();
            Console.WriteLine("=== Encryption Summary ===");
            Console.WriteLine($"Total Excel files found: {totalFiles}");
            Console.WriteLine($"Successfully encrypted: {encryptedCount}");
            Console.WriteLine($"Failed encryptions: {failedCount}");
        }
    }
}
