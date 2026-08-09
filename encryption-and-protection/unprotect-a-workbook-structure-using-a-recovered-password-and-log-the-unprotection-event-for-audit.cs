// Title: Unprotect Excel workbook structure with a recovered password and log the event using Aspose.Cells for .NET
// Description: Loads a password‑protected .xlsx file, calls Workbook.Unprotect with the recovered password (no‑op if already unprotected), writes a UTC timestamped audit entry to the console, and saves the workbook as an unprotected copy.
// Keywords: Aspose.Cells unprotect workbook | C# remove Excel structure protection | recover password Excel file | audit log workbook unprotection | save unprotected workbook Aspose | .NET Excel security | Workbook.Unprotect example
// Common Searches: how to unprotect an Excel workbook with Aspose.Cells | unprotect workbook structure using recovered password C# | log workbook unprotection for compliance | Aspose.Cells example to remove workbook protection | save unprotected Excel file after password removal
// Developer Intent: Remove structure protection from an Excel workbook using a known password and record the action for audit purposes.
// Use Cases: Automated processing of secured Excel files where the password is known or recovered. | Compliance‑driven environments that require a timestamped log whenever protection is removed. | Creating an unprotected copy of a workbook for downstream analysis or distribution.
// AI Prompts: Generate C# code with Aspose.Cells that loads a protected workbook, unprotects it using a given password, logs the operation with a UTC timestamp, and saves the result. | Show how to handle exceptions when loading, unprotecting, and saving an Excel file with Aspose.Cells in .NET. | Explain how to verify workbook protection status before calling Unprotect and how to produce an audit‑ready log entry.

using System;
using System.IO;
using Aspose.Cells;

// Loads a password‑protected .xlsx file, calls Workbook.Unprotect with the recovered password (no‑op if already unprotected), writes a UTC timestamped audit entry to the console, and saves the workbook as an unprotected copy.
class UnprotectWorkbookDemo
{
    static void Main()
    {
        // Path to the protected workbook
        string inputPath = "protected_workbook.xlsx";
        // Path where the unprotected workbook will be saved
        string outputPath = "unprotected_workbook.xlsx";
        // Recovered password used to unprotect the workbook
        string recoveredPassword = "recoveredPassword";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        try
        {
            // Attempt to unprotect using the recovered password.
            // If the workbook is not password‑protected, Unprotect does nothing.
            workbook.Unprotect(recoveredPassword);
            Console.WriteLine($"[{DateTime.UtcNow:u}] Workbook '{inputPath}' was unprotected (or was not protected).");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during unprotecting workbook: {ex.Message}");
            return;
        }

        try
        {
            // Save the (now) unprotected workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Unprotected workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
