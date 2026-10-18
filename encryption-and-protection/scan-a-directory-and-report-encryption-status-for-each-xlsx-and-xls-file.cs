// Title: C# program to scan a folder and report encryption status of XLS and XLSX workbooks using Aspose.Cells
// Common Searches: how to check if an Excel file is password protected with Aspose.Cells in C# | C# code to list encrypted and unencrypted .xls and .xlsx files in a directory | detect encrypted Excel workbooks using Aspose.Cells without providing a password | scan a folder for password‑protected Excel files with Aspose.Cells .NET | determine encryption status of Excel files programmatically in C#
// Tags: Aspose.Cells workbook encryption detection | C# scan directory for Excel file protection | load Excel workbook without password handling | enumerate .xls and .xlsx encryption status | CellsException based encryption check

using System;
using System.IO;
using Aspose.Cells;

// The example iterates through all .xls and .xlsx files in a specified folder (non‑recursive), attempts to open each workbook with Aspose.Cells without a password, treats a CellsException as an indication that the file is encrypted, and writes the file name together with its status (Not Encrypted, Encrypted, or error) to the console.
class EncryptionScanner
{
    static void Main(string[] args)
    {
        // Specify the directory to scan. Change this path as needed.
        string directoryPath = @"C:\Path\To\Your\Folder";

        // Verify that the directory exists.
        if (!Directory.Exists(directoryPath))
        {
            Console.WriteLine($"Directory not found: {directoryPath}");
            return;
        }

        // Get all files in the directory (non‑recursive).
        string[] files = Directory.GetFiles(directoryPath, "*.*", SearchOption.TopDirectoryOnly);

        foreach (string filePath in files)
        {
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension != ".xls" && extension != ".xlsx")
                continue; // Skip non‑Excel files.

            string status;

            // Ensure the file still exists before attempting to load.
            if (!File.Exists(filePath))
            {
                status = "File not found";
                Console.WriteLine($"{Path.GetFileName(filePath)} : {status}");
                continue;
            }

            try
            {
                // Attempt to load the workbook without a password.
                // If the file is password‑protected, Aspose.Cells will throw a CellsException.
                Workbook workbook = new Workbook(filePath);
                // If loading succeeds, the file is not encrypted.
                status = "Not Encrypted";
            }
            catch (CellsException)
            {
                // Assume the file is encrypted when a CellsException occurs during load.
                status = "Encrypted";
            }
            catch (Exception ex)
            {
                // Report any other unexpected errors.
                status = $"Error ({ex.Message})";
            }

            Console.WriteLine($"{Path.GetFileName(filePath)} : {status}");
        }
    }
}
