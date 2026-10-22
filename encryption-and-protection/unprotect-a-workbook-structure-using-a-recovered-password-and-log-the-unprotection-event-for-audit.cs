// Title: How to unprotect an Excel workbook’s structure with a recovered password and log the action using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# unprotect workbook structure with known password | log Excel workbook unprotection timestamp using Aspose.Cells | remove workbook protection and save unprotected file in .NET | verify file exists before unprotecting Excel workbook with Aspose.Cells | audit trail for Excel workbook unprotect operation in C#
// Tags: Aspose.Cells workbook.Unprotect password | C# audit logging Excel unprotection | Excel workbook structure protection removal Aspose.Cells | save unprotected workbook Aspose.Cells C# | file existence check before loading workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the Excel file’s existence, loads it with Aspose.Cells, removes structure and window protection using the recovered password via Workbook.Unprotect, records a UTC timestamped audit log, saves the unprotected workbook to a new file, and handles any errors with detailed logging.
class WorkbookUnprotector
{
    static void Main()
    {
        // Path to the workbook to be unprotected
        string inputPath = "ProtectedWorkbook.xlsx";
        // Path where the unprotected workbook will be saved
        string outputPath = "UnprotectedWorkbook.xlsx";

        // Recovered password for the workbook structure protection
        string recoveredPassword = "MyRecoveredPassword";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"[{DateTime.UtcNow:u}] ERROR: Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Unprotect the workbook (structure and windows) using the recovered password
            // Aspose.Cells provides a direct Unprotect method on Workbook for this purpose
            workbook.Unprotect(recoveredPassword);

            // Log the successful unprotection event for audit
            Console.WriteLine($"[{DateTime.UtcNow:u}] Workbook unprotected successfully. File: {inputPath}");

            // Save the unprotected workbook
            workbook.Save(outputPath);
            Console.WriteLine($"[{DateTime.UtcNow:u}] Unprotected workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any errors that occur during the process
            Console.WriteLine($"[{DateTime.UtcNow:u}] ERROR: Failed to unprotect workbook. Details: {ex.Message}");
        }
    }
}
