// Title: How to unprotect an Excel workbook’s structure with a recovered password and log the action using Aspose.Cells for .NET
// Common Searches: how to remove workbook structure protection using Aspose.Cells in C# | c# log workbook unprotection event for compliance with Aspose.Cells | unprotect an Excel file with a known password and save it using Aspose.Cells .NET | asp.net check if protected workbook file exists before calling Workbook.Unprotect | aspose.cells write audit entry to console after unprotecting workbook
// Tags: Aspose.Cells workbook.unprotect method | unprotect Excel workbook structure Aspose.Cells | audit logging of workbook unprotection .NET | save unprotected Excel file Aspose.Cells | validate workbook file existence Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads a password‑protected Excel workbook, uses Workbook.Unprotect with the recovered password to remove structure and window protection, writes a UTC timestamped audit message to the console, and saves the workbook as an unprotected file while handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        // Path to the protected workbook
        string inputPath = "protected.xlsx";

        // Recovered password for the workbook structure
        string password = "RecoveredPassword123";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Unprotect the workbook (structure and windows) using the recovered password
            workbook.Unprotect(password);

            // Log the unprotection event for audit purposes
            Console.WriteLine($"[{DateTime.UtcNow:u}] Workbook unprotected for file '{inputPath}' using password '{password}'.");

            // Save the workbook after unprotection (overwrite or save as new file)
            string outputPath = "unprotected.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
