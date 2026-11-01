// Title: Detect and log password‑protected Excel workbooks in a directory using Aspose.Cells for .NET (C#)
// Common Searches: C# code to identify encrypted .xlsx files with Aspose.Cells | How to programmatically check if an Excel workbook is password protected using Aspose.Cells | Scan a folder for password‑protected Excel files and output their status in .NET | Aspose.Cells detect password protection without opening workbook C# | List encrypted Excel spreadsheets in a directory using Aspose.Cells API
// Tags: detect encrypted Excel workbook Aspose.Cells | enumerate password protected .xlsx files C# | log encryption status of spreadsheets .NET | handle CellsException for protected workbooks | scan directory for Excel files Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program walks through a specified folder, filters for common Excel extensions, attempts to load each workbook with Aspose.Cells without a password, catches CellsException to determine if the file is password‑protected, and writes the file name together with a true/false encryption flag to the console while handling other I/O errors.
class EncryptedExcelScanner
{
    static void Main(string[] args)
    {
        // Specify the directory to scan. Change this path as needed.
        string directoryPath = @"C:\ExcelFiles";

        // Verify the directory exists.
        if (!Directory.Exists(directoryPath))
        {
            Console.WriteLine($"Directory not found: {directoryPath}");
            return;
        }

        // Get all files with Excel extensions in the directory (including subfolders if desired).
        string[] excelFiles = Directory.GetFiles(directoryPath, "*.*", SearchOption.TopDirectoryOnly);
        foreach (string filePath in excelFiles)
        {
            // Filter by typical Excel extensions.
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension != ".xls" && extension != ".xlsx" && extension != ".xlsm" && extension != ".xlsb")
                continue;

            bool isEncrypted = false;

            try
            {
                // Attempt to load the workbook without providing a password.
                // If the file is encrypted, Aspose.Cells will throw an exception.
                Workbook workbook = new Workbook(filePath);
                // If load succeeds, the file is not encrypted.
                isEncrypted = false;
            }
            catch (CellsException ex)
            {
                // Aspose.Cells throws CellsException for password‑protected files.
                // Check the exception message to confirm encryption status.
                if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    isEncrypted = true;
                }
                else
                {
                    // For other exceptions, rethrow or handle as needed.
                    Console.WriteLine($"Error processing file '{Path.GetFileName(filePath)}': {ex.Message}");
                    continue;
                }
            }
            catch (Exception ex)
            {
                // Handle any non‑Aspose exceptions (e.g., file access issues).
                Console.WriteLine($"Unexpected error with file '{Path.GetFileName(filePath)}': {ex.Message}");
                continue;
            }

            // Log the result.
            Console.WriteLine($"{Path.GetFileName(filePath)} - Encrypted: {isEncrypted}");
        }
    }
}
