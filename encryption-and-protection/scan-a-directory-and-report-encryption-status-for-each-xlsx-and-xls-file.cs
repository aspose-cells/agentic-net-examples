// Title: Identify encrypted XLSX and XLS workbooks in a folder using Aspose.Cells for .NET
// AI Prompts: Write C# code that scans a given directory, attempts to load each .xlsx or .xls file with Aspose.Cells LoadOptions, and marks the file as encrypted when a CellsException is thrown. | Create a reusable method that returns a dictionary mapping Excel file paths to a boolean indicating password‑protected status, using Aspose.Cells without providing a password. | Enhance the program to log files that cannot be opened due to corruption or unsupported format while still reporting their encryption status.
// Common Searches: C# how to check if an Excel file is password protected with Aspose.Cells | list encrypted .xlsx files in a directory using Aspose.Cells .NET | detect workbook encryption without password using Aspose.Cells LoadOptions | scan folder for protected Excel workbooks and output status in C# | handle CellsException to identify protected spreadsheets in .NET
// Tags: detect encrypted Excel workbooks Aspose.Cells | folder scan for .xlsx encryption status .NET | load workbook without password Aspose.Cells | handle CellsException for password‑protected spreadsheets | enumerate Excel file protection using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Scans a specified directory, attempts to load each .xlsx or .xls file with Aspose.Cells, catches CellsException to determine if the workbook is password‑protected, and prints the file name with its encryption status while optionally logging corrupt or unsupported files.
class Program
{
    static void Main()
    {
        // Set the directory to scan
        string folderPath = @"C:\Path\To\Directory"; // TODO: change to your folder

        // Verify that the directory exists
        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Directory not found: {folderPath}");
            return;
        }

        // Get all files in the directory (top level only)
        string[] files = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly);

        foreach (string filePath in files)
        {
            string extension = Path.GetExtension(filePath).ToLowerInvariant();

            // Process only XLSX and XLS files
            if (extension == ".xlsx" || extension == ".xls")
            {
                // Ensure the file actually exists before attempting to load
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                bool isEncrypted = false;

                try
                {
                    // Attempt to load the workbook without a password
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
                    Workbook workbook = new Workbook(filePath, loadOptions);

                    // If loading succeeds, the file is not encrypted
                    isEncrypted = false;
                }
                catch (CellsException)
                {
                    // Aspose.Cells throws CellsException when the file is password protected
                    isEncrypted = true;
                }
                catch (Exception ex)
                {
                    // Other unexpected errors (corrupt file, unsupported format, etc.)
                    Console.WriteLine($"Error processing {Path.GetFileName(filePath)}: {ex.Message}");
                    continue;
                }

                // Output the encryption status
                Console.WriteLine($"{Path.GetFileName(filePath)} : {(isEncrypted ? "Encrypted" : "Not Encrypted")}");
            }
        }
    }
}
