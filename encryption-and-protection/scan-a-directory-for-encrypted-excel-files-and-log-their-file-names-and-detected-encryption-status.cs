// Title: C# – Scan a directory for encrypted Excel files using Aspose.Cells
// Description: A console utility that recursively scans a folder, filters Excel workbooks, uses Aspose.Cells.FileFormatUtil to detect password protection, and logs each file name with its encryption status or any error encountered.
// Keywords: Aspose.Cells | C# | encrypted Excel files | password protected workbook | FileFormatUtil | detect Excel encryption | scan folder for Excel | batch Excel security check | Excel file protection .NET | GitHub example
// Common Searches: Aspose.Cells detect encrypted workbook C# | C# code to list password protected Excel files | How to check if an Excel file is encrypted using Aspose | Scan folder for protected Excel files .NET | Batch encryption detection Aspose.Cells
// Developer Intent: Determine which Excel files in a given directory are password‑protected and output their names with a true/false encryption flag.
// Use Cases: Generate an inventory of encrypted workbooks before a migration or backup. | Exclude password‑protected files from bulk conversion, data extraction, or analytics pipelines. | Alert administrators to unexpected encryption or corrupted files during nightly scans. | Perform compliance audits of confidential spreadsheets across shared drives.
// AI Prompts: Create a method that accepts a folder path and returns a list of (file name, isEncrypted) tuples using Aspose.Cells. | Rewrite the program to write the results to a CSV file with columns: FileName, Encrypted, ErrorMessage. | Add timestamped file logging and preserve console output for each detection attempt. | Implement parallel processing to speed up scanning of large file collections. | Extend the example to output results as JSON for consumption by a REST API.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionScanner
{
    // A console utility that recursively scans a folder, filters Excel workbooks, uses Aspose.Cells.FileFormatUtil to detect password protection, and logs each file name with its encryption status or any error encountered.
    class Program
    {
        static void Main(string[] args)
        {
            // Directory to scan – change as needed or pass as an argument
            string folderPath = args.Length > 0 ? args[0] : @"C:\ExcelFiles";

            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder does not exist: {folderPath}");
                return;
            }

            // Define Excel file extensions to consider
            string[] excelExtensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb", ".xls", ".xlt", ".xltx", ".xltm" };

            // Enumerate files recursively
            foreach (string filePath in Directory.EnumerateFiles(folderPath, "*.*", SearchOption.AllDirectories))
            {
                // Skip non‑Excel files
                if (Array.IndexOf(excelExtensions, Path.GetExtension(filePath).ToLowerInvariant()) < 0)
                    continue;

                try
                {
                    // Detect file format and encryption status
                    FileFormatInfo info = FileFormatUtil.DetectFileFormat(filePath);
                    bool isEncrypted = info.IsEncrypted;

                    // Log result
                    Console.WriteLine($"{Path.GetFileName(filePath)}\tEncrypted: {isEncrypted}");
                }
                catch (Exception ex)
                {
                    // Log any errors (e.g., corrupted file)
                    Console.WriteLine($"{Path.GetFileName(filePath)}\tError: {ex.Message}");
                }
            }
        }
    }
}
