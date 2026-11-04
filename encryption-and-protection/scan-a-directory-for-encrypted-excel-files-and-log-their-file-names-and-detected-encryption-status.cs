// Title: C# program to scan a folder and log encryption status of Excel files using Aspose.Cells
// Common Searches: C# scan folder for password‑protected Excel files using Aspose.Cells | How to determine if an .xlsx workbook is encrypted with Aspose.Cells in .NET | List encrypted Excel files in a directory with Aspose.Cells LoadOptions | Detect password protected Excel workbooks without opening them in C# | Log encryption status of multiple Excel files using Aspose.Cells
// Tags: folder enumeration of encrypted Excel workbooks Aspose.Cells | password detection for Excel files C# Aspose.Cells | load options for checking workbook encryption Aspose.Cells | console logging of Excel encryption status .NET | exception handling for protected Excel workbooks Aspose.Cells

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;

// The C# utility scans a specified directory for Excel files (.xls, .xlsx, .xlsm, .xlsb), attempts to load each workbook with Aspose.Cells without a password, catches CellsException messages that indicate password protection, and writes the file name together with a true/false encryption flag to the console while gracefully handling other errors.
class Program
{
    static void Main(string[] args)
    {
        // Determine folder to scan: argument or default path.
        string folderPath = args.Length > 0 ? args[0] : @"C:\ExcelFiles";

        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder not found: {folderPath}");
            return;
        }

        // Supported Excel extensions.
        string[] extensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb" };

        // Retrieve files with the specified extensions (non‑recursive).
        var excelFiles = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly)
                                  .Where(f => extensions.Contains(Path.GetExtension(f), StringComparer.OrdinalIgnoreCase));

        foreach (var filePath in excelFiles)
        {
            bool isEncrypted = false;

            try
            {
                // Verify the file exists before attempting to load.
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                // Try to load the workbook without a password.
                var loadOptions = new LoadOptions(LoadFormat.Auto);
                var workbook = new Workbook(filePath, loadOptions);
                // If loading succeeds, the file is not encrypted.
                isEncrypted = false;
            }
            catch (CellsException ex)
            {
                // Aspose.Cells throws CellsException for password‑protected files.
                // Check the message to determine if it is due to a password.
                if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    isEncrypted = true;
                }
                else
                {
                    // Log other CellsException issues and skip the file.
                    Console.WriteLine($"CellsException for '{Path.GetFileName(filePath)}': {ex.Message}");
                    continue;
                }
            }
            catch (Exception ex)
            {
                // Log unexpected errors (e.g., corrupted file) and skip to next file.
                Console.WriteLine($"Error checking file '{Path.GetFileName(filePath)}': {ex.Message}");
                continue;
            }

            // Output the result.
            Console.WriteLine($"{Path.GetFileName(filePath)} - Encrypted: {isEncrypted}");
        }
    }
}
