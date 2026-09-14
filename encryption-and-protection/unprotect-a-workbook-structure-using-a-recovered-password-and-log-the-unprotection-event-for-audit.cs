// Title: Unprotect an Excel workbook’s structure with a recovered password and record the audit log using Aspose.Cells for .NET
// AI Prompts: Load a password‑protected .xlsx file with Aspose.Cells, call Workbook.Unprotect with the recovered password, and save the result to a new workbook. | Insert code that writes an audit entry containing the UTC timestamp, file name, and unprotection action to the console or a log file after the workbook is unprotected. | Refactor the sample to accept input and output file paths as parameters and ensure the original protected workbook remains unchanged.
// Common Searches: asp.net unprotect workbook structure using Aspose.Cells with known password | c# log excel workbook unprotection timestamp for audit | how to remove workbook windows protection programmatically with Aspose.Cells | save unprotected Excel file to a different folder using Aspose.Cells .NET
// Tags: Aspose.Cells workbook.Unprotect method C# | remove workbook structure protection programmatically | audit logging for Excel workbook unprotection | save unprotected Excel workbook Aspose.Cells | recover password for Excel protection .NET

using System;
using System.IO;
using Aspose.Cells;

// The program loads a protected Excel workbook, uses a recovered password to unprotect its structure and windows via Aspose.Cells, logs the unprotection event with a UTC timestamp for audit, and saves the unprotected workbook to a new file.
class WorkbookUnprotector
{
    static void Main()
    {
        // Path to the workbook to be unprotected
        string inputPath = "ProtectedWorkbook.xlsx";

        // Recovered password for the workbook structure
        string recoveredPassword = "MyRecoveredPassword";

        // Verify the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Unprotect the workbook (structure and windows) using the recovered password
            workbook.Unprotect(recoveredPassword);

            // Log the unprotection event for audit purposes
            Console.WriteLine($"[{DateTime.UtcNow:u}] Workbook structure unprotected for file '{inputPath}' using recovered password.");

            // Save the workbook (optionally to a new file to preserve the original)
            string outputPath = "UnprotectedWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
