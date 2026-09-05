// Title: C# program to scan a directory for password‑protected Excel files and log each file’s encryption status with Aspose.Cells
// AI Prompts: Create a C# console app that enumerates all .xls, .xlsx, .xlsm, and .xlsb files in a specified folder, attempts to open each workbook using Aspose.Cells, and prints ‘Encrypted’ or ‘Not Encrypted’ for every file. | Update the scanner to write the file name, encryption result, and any error messages to a CSV report instead of the console. | Extend the solution to recursively search subfolders while preserving the detection of password‑protected workbooks. | Add an option to supply a list of possible passwords and try to open encrypted workbooks, logging whether the correct password was found.
// Common Searches: how to detect password protected Excel workbooks using Aspose.Cells in C# | C# scan folder for encrypted .xlsx files Aspose.Cells example | list all encrypted Excel files in a directory with Aspose.Cells | catch CellsException for password required when loading a workbook in C# | generate CSV report of Excel file encryption status using Aspose.Cells
// Tags: detect encrypted Excel files Aspose.Cells | C# directory scan for .xls .xlsx .xlsm .xlsb | log workbook encryption status console | handle CellsException password prompt | recursive folder traversal Aspose.Cells | export encryption results to CSV C#

using System;
using System.IO;
using Aspose.Cells;

// The sample iterates through a given folder, filters for common Excel extensions, attempts to load each workbook with Aspose.Cells, and uses the presence of a password‑related CellsException to determine if the file is encrypted. It logs each file name with either "Encrypted", "Not Encrypted", or an error description, providing a basis for further reporting or automation.
class ExcelEncryptionScanner
{
    static void Main()
    {
        // Specify the directory to scan
        string folderPath = @"C:\Path\To\Directory";

        // Verify the directory exists
        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Directory not found: {folderPath}");
            return;
        }

        // Get all files in the directory (filter later by extension)
        string[] files = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly);

        foreach (string filePath in files)
        {
            // Filter by typical Excel extensions
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension != ".xls" && extension != ".xlsx" && extension != ".xlsm" && extension != ".xlsb")
                continue;

            // Ensure the file still exists before attempting to load
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"{Path.GetFileName(filePath)} - File not found");
                continue;
            }

            try
            {
                // Attempt to load the workbook without a password
                Workbook workbook = new Workbook(filePath);
                // If loading succeeds, the file is not encrypted
                Console.WriteLine($"{Path.GetFileName(filePath)} - Not Encrypted");
            }
            catch (CellsException ex)
            {
                // Aspose.Cells throws CellsException when a password is required
                if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Console.WriteLine($"{Path.GetFileName(filePath)} - Encrypted");
                }
                else
                {
                    // Other CellsException (e.g., corrupted file)
                    Console.WriteLine($"{Path.GetFileName(filePath)} - CellsException: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                // Log any other unexpected errors
                Console.WriteLine($"{Path.GetFileName(filePath)} - Error: {ex.Message}");
            }
        }
    }
}
