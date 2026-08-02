// Title: C# – Scan a directory for encrypted Excel files using Aspose.Cells FileFormatUtil
// Description: A console utility that enumerates *.xls, *.xlsx, *.xlsm and *.xlsb files in a given folder, uses Aspose.Cells FileFormatUtil.DetectFileFormat to determine whether each workbook is password‑protected, and writes the file name with a true/false encryption flag to the console while handling I/O and format errors.
// Keywords: Aspose.Cells encryption detection C# | FileFormatUtil DetectFileFormat encrypted workbook | scan folder for password‑protected Excel files | list encrypted .xlsx files .NET | check Excel file protection programmatically
// Common Searches: how to detect encrypted Excel files with Aspose.Cells | C# code to list password‑protected workbooks in a folder | using FileFormatUtil to identify protected spreadsheets | batch scan for encrypted .xls/.xlsx files | Aspose.Cells example for encryption status
// Developer Intent: Find all Excel workbooks in a specified directory and report which ones are encrypted.
// Use Cases: Generate a compliance inventory of password‑protected spreadsheets before migration. | Skip encrypted files automatically during a bulk data‑import operation. | Audit a shared drive to quantify protected Excel files for security reporting.
// AI Prompts: Create a C# script that scans subfolders recursively, detects encrypted Excel files with Aspose.Cells, and exports the results to a CSV file. | Show how to catch specific exceptions from FileFormatUtil.DetectFileFormat when a file is corrupted or uses an unsupported format. | Refactor the sample to use parallel processing for faster scanning of large directories and summarize the total count of encrypted versus unencrypted files.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionScanner
{
    // A console utility that enumerates *.xls, *.xlsx, *.xlsm and *.xlsb files in a given folder, uses Aspose.Cells FileFormatUtil.DetectFileFormat to determine whether each workbook is password‑protected, and writes the file name with a true/false encryption flag to the console while handling I/O and format errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Directory to scan – can be overridden by a command‑line argument
            string folderPath = args.Length > 0 ? args[0] : @"C:\Path\To\Excel\Files";

            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder not found: {folderPath}");
                return;
            }

            // Supported Excel extensions
            string[] extensions = new[] { "*.xls", "*.xlsx", "*.xlsm", "*.xlsb" };

            try
            {
                // Iterate through each extension and process matching files
                foreach (string ext in extensions)
                {
                    foreach (string filePath in Directory.GetFiles(folderPath, ext, SearchOption.TopDirectoryOnly))
                    {
                        try
                        {
                            // Detect file format and encryption status
                            FileFormatInfo info = FileFormatUtil.DetectFileFormat(filePath);

                            // Log file name and encryption status
                            Console.WriteLine($"File: {Path.GetFileName(filePath)} | Encrypted: {info.IsEncrypted}");
                        }
                        catch (Exception ex)
                        {
                            // Log any errors encountered while processing the file
                            Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Log unexpected errors during directory enumeration
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
