// Title: Create a C# console utility with Aspose.Cells to batch encrypt Excel files in a directory and generate a processing summary
// Common Searches: how to batch encrypt .xlsx files with Aspose.Cells in a C# console app | C# program to add password protection to all Excel workbooks in a folder | using Aspose.Cells to encrypt multiple Excel files and output a summary report | recursive directory scan for Excel files and apply workbook password in .NET | command line tool for encrypting .xls .xlsx .xlsm files with Aspose.Cells
// Tags: Aspose.Cells batch workbook encryption | C# console Excel password protection | encrypt .xls .xlsx .xlsm files programmatically | recursive file system scan for Excel workbooks | encryption summary reporting in .NET

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

namespace ExcelEncryptor
{
    // A C# console application that recursively scans a given directory for .xls, .xlsx, and .xlsm files, loads each workbook with Aspose.Cells, applies a fixed password for encryption, saves the file in its original format, and prints a summary of total files processed, successfully encrypted, and failures.
    class Program
    {
        // Password used for encrypting all Excel files.
        private const string EncryptionPassword = "StrongPassword123";

        static void Main(string[] args)
        {
            try
            {
                // Validate command‑line arguments.
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

                // Supported Excel extensions.
                string[] extensions = new[] { ".xls", ".xlsx", ".xlsm" };

                // Collect all Excel files recursively.
                var excelFiles = Directory.GetFiles(directoryPath, "*.*", SearchOption.AllDirectories);
                var targetFiles = new List<string>();

                foreach (var file in excelFiles)
                {
                    if (Array.Exists(extensions, ext => ext.Equals(Path.GetExtension(file), StringComparison.OrdinalIgnoreCase)))
                    {
                        targetFiles.Add(file);
                    }
                }

                int totalFiles = targetFiles.Count;
                int successCount = 0;
                int failureCount = 0;

                Console.WriteLine($"Found {totalFiles} Excel file(s) to encrypt.");

                foreach (var filePath in targetFiles)
                {
                    try
                    {
                        // Ensure the file still exists before loading.
                        if (!File.Exists(filePath))
                        {
                            Console.WriteLine($"Skipped (not found): {filePath}");
                            failureCount++;
                            continue;
                        }

                        // Load the workbook.
                        Workbook workbook = new Workbook(filePath);

                        // Set password for encryption.
                        workbook.Settings.Password = EncryptionPassword;

                        // Determine original extension and save with appropriate format.
                        string ext = Path.GetExtension(filePath).ToLowerInvariant();
                        switch (ext)
                        {
                            case ".xls":
                                workbook.Save(filePath, SaveFormat.Excel97To2003);
                                break;

                            case ".xlsx":
                                workbook.Save(filePath, SaveFormat.Xlsx);
                                break;

                            case ".xlsm":
                                workbook.Save(filePath, SaveFormat.Xlsm);
                                break;

                            default:
                                // Fallback to Xlsx if an unexpected extension appears.
                                workbook.Save(filePath, SaveFormat.Xlsx);
                                break;
                        }

                        successCount++;
                        Console.WriteLine($"Encrypted: {filePath}");
                    }
                    catch (Exception ex)
                    {
                        failureCount++;
                        Console.WriteLine($"Failed to encrypt \"{filePath}\": {ex.Message}");
                    }
                }

                // Summary report.
                Console.WriteLine();
                Console.WriteLine("Encryption Summary:");
                Console.WriteLine($"Total files processed : {totalFiles}");
                Console.WriteLine($"Successfully encrypted: {successCount}");
                Console.WriteLine($"Failed encryptions    : {failureCount}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
