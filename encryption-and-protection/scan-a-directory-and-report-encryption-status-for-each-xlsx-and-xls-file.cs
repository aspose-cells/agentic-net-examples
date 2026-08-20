// Title: Detect Encryption of .xlsx and .xls Files in a Folder Using Aspose.Cells for .NET (C#)
// Description: A console utility that scans a specified directory, filters Excel workbooks (.xlsx, .xls), and uses Aspose.Cells FileFormatUtil.DetectFileFormat to report whether each file is encrypted. It outputs the file name with a true/false flag or an error message for unsupported or corrupted files.
// Keywords: Aspose.Cells C# encryption detection | FileFormatUtil IsEncrypted | detect password protected Excel | scan folder for encrypted .xlsx | list encrypted Excel files .NET | Excel file encryption status | Aspose.Cells DetectFileFormat
// Common Searches: C# check if Excel file is password protected Aspose.Cells | how to list encrypted .xls files in a directory using Aspose | detect encrypted workbook with Aspose.Cells .NET | FileFormatUtil DetectFileFormat encryption flag example | scan folder for encrypted Excel workbooks C#
// Developer Intent: Determine the encryption (password‑protection) state of each Excel workbook in a given folder.
// Use Cases: Generate a compliance report that separates encrypted from unencrypted workbooks before bulk processing. | Skip or log encrypted files in an automated conversion pipeline that only handles unprotected Excel files. | Capture detection errors for corrupted or unsupported Excel files while scanning a directory. | Perform a security audit of shared drives to identify password‑protected spreadsheets.
// AI Prompts: Create a C# method that returns a Dictionary<string, bool> mapping Excel file names to their encryption status using Aspose.Cells. | Extend the sample to recursively scan subfolders and export results (file name, encrypted flag, error message) to a CSV file. | Provide best‑practice error handling for FileFormatUtil.DetectFileFormat when processing large batches of Excel files. | Write unit tests that verify encryption detection for both .xlsx and .xls files with and without passwords.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionScanner
{
    // A console utility that scans a specified directory, filters Excel workbooks (.xlsx, .xls), and uses Aspose.Cells FileFormatUtil.DetectFileFormat to report whether each file is encrypted. It outputs the file name with a true/false flag or an error message for unsupported or corrupted files.
    class Program
    {
        static void Main(string[] args)
        {
            // Directory to scan – change as needed or pass as an argument
            string folderPath = args.Length > 0 ? args[0] : Directory.GetCurrentDirectory();

            // Get all .xlsx and .xls files in the directory (non‑recursive)
            string[] excelFiles = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly);
            foreach (string filePath in excelFiles)
            {
                string extension = Path.GetExtension(filePath).ToLowerInvariant();
                if (extension != ".xlsx" && extension != ".xls")
                    continue; // Skip non‑Excel files

                try
                {
                    // Detect file format and encryption status using Aspose.Cells API
                    FileFormatInfo info = FileFormatUtil.DetectFileFormat(filePath);
                    bool isEncrypted = info.IsEncrypted;

                    Console.WriteLine($"{Path.GetFileName(filePath)} - Encrypted: {isEncrypted}");
                }
                catch (Exception ex)
                {
                    // If detection fails, report the error but continue processing other files
                    Console.WriteLine($"{Path.GetFileName(filePath)} - Error detecting encryption: {ex.Message}");
                }
            }
        }
    }
}
