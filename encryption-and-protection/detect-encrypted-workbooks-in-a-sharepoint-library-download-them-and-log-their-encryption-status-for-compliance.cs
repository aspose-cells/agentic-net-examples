// Title: C# Sample: Detect and Log Encrypted Excel Workbooks Using Aspose.Cells (Local Folder – Ready for SharePoint)
// Description: A console app that scans a specified directory, filters Excel files, and uses Aspose.Cells FileFormatUtil to identify encrypted workbooks. It prints each file name with its encryption flag and includes robust error handling. The code can be extended to download files from a SharePoint library before checking compliance.
// Keywords: Aspose.Cells encryption detection C# | FileFormatUtil IsEncrypted example | check Excel password protection .NET | scan folder for protected Excel files | audit encrypted workbooks | SharePoint Excel encryption check | US compliance Excel security | EU GDPR Excel encryption
// Common Searches: how to detect password‑protected Excel files with Aspose.Cells | C# code to list encrypted .xlsx files in a directory | Aspose.Cells FileFormatUtil IsEncrypted usage | log encryption status of multiple Excel workbooks | download and verify encrypted Excel files from SharePoint
// Developer Intent: Find out which Excel workbooks are encrypted and output their status for compliance reporting.
// Use Cases: Run a compliance scan on a repository of Excel files before publishing. | Generate an audit report showing encrypted vs. unencrypted workbooks. | Integrate encryption detection into an automated pipeline that moves protected files to a secure location.
// AI Prompts: Write a C# method that returns true if a given Excel file is encrypted using Aspose.Cells. | Show how to modify the sample to export the file name and encryption flag to a CSV file. | Explain how to catch Aspose.Cells specific exceptions during encryption detection and log detailed error information.

using System;
using System.IO;
using Aspose.Cells; // Requires Aspose.Cells for .NET

namespace SharePointWorkbookEncryptionCheck
{
    // A console app that scans a specified directory, filters Excel files, and uses Aspose.Cells FileFormatUtil to identify encrypted workbooks. It prints each file name with its encryption flag and includes robust error handling. The code can be extended to download files from a SharePoint library before checking compliance.
    class Program
    {
        // Adjust this path to point to a folder that contains Excel files to be checked
        private const string LocalFolderPath = @"C:\Temp\ExcelFiles";

        static void Main()
        {
            try
            {
                // Verify that the folder exists
                if (!Directory.Exists(LocalFolderPath))
                {
                    Console.WriteLine($"Folder not found: {LocalFolderPath}");
                    return;
                }

                // Get all Excel files in the folder (including subfolders if needed)
                string[] excelFiles = Directory.GetFiles(LocalFolderPath, "*.*", SearchOption.TopDirectoryOnly);
                if (excelFiles.Length == 0)
                {
                    Console.WriteLine("No files found in the specified folder.");
                    return;
                }

                foreach (string filePath in excelFiles)
                {
                    // Filter only supported Excel extensions
                    string extension = Path.GetExtension(filePath).ToLowerInvariant();
                    if (extension != ".xls" && extension != ".xlsx" && extension != ".xlsm" && extension != ".xlsb")
                    {
                        continue;
                    }

                    // Ensure the file actually exists before processing
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found: {filePath}");
                        continue;
                    }

                    bool isEncrypted = false;
                    try
                    {
                        // Detect file format and encryption status using Aspose.Cells
                        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
                        isEncrypted = formatInfo.IsEncrypted;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error detecting format for '{Path.GetFileName(filePath)}': {ex.Message}");
                    }

                    // Log the result
                    Console.WriteLine($"File: {Path.GetFileName(filePath)} | Encrypted: {isEncrypted}");
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors to prevent the application from crashing
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
