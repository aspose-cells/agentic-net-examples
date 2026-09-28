// Title: Create a C# command‑line tool that recursively encrypts all Excel workbooks in a folder using Aspose.Cells and prints an encryption summary
// AI Prompts: Write a C# console program that takes a directory path, recursively finds .xls, .xlsx, .xlsm, and .xlsb files, applies a password with Aspose.Cells Workbook.Settings.Password, saves each file using the proper SaveOptions, and displays counts of total, successful, and failed encryptions. | Add support for an optional second command‑line argument that lets the user specify the encryption password, defaulting to a hard‑coded value when omitted. | Update the utility so that encrypted copies are written to a separate output directory while preserving the original folder hierarchy, leaving the source files untouched.
// Common Searches: how to encrypt multiple Excel files with Aspose.Cells in a .NET console app | C# batch password protect .xlsx files from command line | recursive folder scan encrypt Excel workbooks Aspose.Cells example | generate summary report of encrypted Excel files using Aspose.Cells C# | save encrypted Excel workbook with correct format using Aspose.Cells SaveOptions
// Tags: encrypt excel workbooks Aspose.Cells C# | recursive directory scan for Excel files | set workbook password Aspose.Cells | saveoptions per excel format Aspose.Cells | command line excel encryption utility

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

namespace ExcelEncryptor
{
    // A C# console utility that accepts a folder path, locates all .xls, .xlsx, .xlsm, and .xlsb workbooks, encrypts each using a fixed password via Aspose.Cells (applying the appropriate SaveOptions for the file type), overwrites the originals, and prints a summary of total files found, successfully encrypted, and any errors.
    class Program
    {
        // Fixed password for encryption – modify as needed
        private const string EncryptionPassword = "Password123";

        static void Main(string[] args)
        {
            // Validate command‑line arguments
            if (args.Length != 1)
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

            // Collect all Excel files (xls, xlsx, xlsm, xlsb)
            string[] excelFiles = Directory.GetFiles(directoryPath, "*.*", SearchOption.AllDirectories);
            List<string> targetFiles = new List<string>();
            foreach (var file in excelFiles)
            {
                string ext = Path.GetExtension(file).ToLowerInvariant();
                if (ext == ".xls" || ext == ".xlsx" || ext == ".xlsm" || ext == ".xlsb")
                {
                    targetFiles.Add(file);
                }
            }

            int totalFiles = targetFiles.Count;
            int encryptedCount = 0;
            int errorCount = 0;

            foreach (var filePath in targetFiles)
            {
                try
                {
                    // Ensure the file exists before attempting to load
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found: {filePath}");
                        errorCount++;
                        continue;
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Set the password for opening the workbook
                    workbook.Settings.Password = EncryptionPassword;

                    // Determine the appropriate SaveOptions based on extension
                    string ext = Path.GetExtension(filePath).ToLowerInvariant();
                    SaveOptions saveOptions;

                    if (ext == ".xls")
                    {
                        saveOptions = new XlsSaveOptions(SaveFormat.Excel97To2003);
                    }
                    else if (ext == ".xlsm")
                    {
                        var opts = new OoxmlSaveOptions(SaveFormat.Xlsm);
                        saveOptions = opts;
                    }
                    else if (ext == ".xlsb")
                    {
                        // Use parameter‑less constructor as the overload with SaveFormat is obsolete
                        saveOptions = new XlsbSaveOptions();
                    }
                    else // .xlsx and any other default
                    {
                        var opts = new OoxmlSaveOptions(SaveFormat.Xlsx);
                        saveOptions = opts;
                    }

                    // Save the workbook with password protection (overwrites original file)
                    workbook.Save(filePath, saveOptions);
                    encryptedCount++;
                    Console.WriteLine($"Encrypted: {filePath}");
                }
                catch (Exception ex)
                {
                    errorCount++;
                    Console.WriteLine($"Error encrypting \"{filePath}\": {ex.Message}");
                }
            }

            // Summary report
            Console.WriteLine();
            Console.WriteLine("=== Encryption Summary ===");
            Console.WriteLine($"Total Excel files found : {totalFiles}");
            Console.WriteLine($"Successfully encrypted   : {encryptedCount}");
            Console.WriteLine($"Failed encryptions       : {errorCount}");
        }
    }
}
