// Title: Scan a directory and report encryption status of XLS and XLSX files using Aspose.Cells for .NET
// Common Searches: C# code to list Excel files and show which are password protected with Aspose.Cells | How to programmatically check if an .xlsx workbook is encrypted using Aspose.Cells in .NET | Detect encrypted Excel files in a folder with Aspose.Cells without opening each file | Aspose.Cells C# detect password‑protected XLS files during batch processing | Identify protected Excel workbooks in a directory using Aspose.Cells exception handling
// Tags: encrypted workbook detection Aspose.Cells | batch Excel protection scan .NET | CellsException password protection handling | XLSX encryption status enumeration | folder level Excel security check

using System;
using System.IO;
using Aspose.Cells;

// The example iterates over all .xls and .xlsx files in a specified folder (non‑recursive), attempts to load each workbook with Aspose.Cells, and treats a CellsException containing the word "password" as an indication that the file is encrypted; otherwise it reports the file as not encrypted. It outputs each file name with its encryption status while handling missing directories, missing files, and other errors.
class Program
{
    static void Main(string[] args)
    {
        // Specify the directory to scan. Change this path as needed.
        string directoryPath = @"C:\Path\To\Your\Folder";

        if (!Directory.Exists(directoryPath))
        {
            Console.WriteLine($"Directory does not exist: {directoryPath}");
            return;
        }

        // Get all files in the directory (non‑recursive).
        string[] files = Directory.GetFiles(directoryPath, "*.*", SearchOption.TopDirectoryOnly);

        foreach (string filePath in files)
        {
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension != ".xls" && extension != ".xlsx")
                continue; // Skip non‑Excel files

            // Ensure the file actually exists before attempting to load.
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                continue;
            }

            bool isEncrypted = false;

            try
            {
                // Attempt to load the workbook without a password.
                // If the file is password‑protected, Aspose.Cells throws an exception.
                var wb = new Workbook(filePath);
                // Loaded successfully => not encrypted.
                isEncrypted = false;
            }
            catch (CellsException ce)
            {
                // Aspose.Cells throws CellsException for password‑protected files.
                // Check the exception message for password‑related text.
                if (ce.Message != null && ce.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    isEncrypted = true;
                }
                else
                {
                    Console.WriteLine($"Error processing '{Path.GetFileName(filePath)}': {ce.Message}");
                    continue;
                }
            }
            catch (Exception ex)
            {
                // Unexpected error while processing the file.
                Console.WriteLine($"Error processing '{Path.GetFileName(filePath)}': {ex.Message}");
                continue;
            }

            string status = isEncrypted ? "Encrypted" : "Not Encrypted";
            Console.WriteLine($"{Path.GetFileName(filePath)} : {status}");
        }
    }
}
