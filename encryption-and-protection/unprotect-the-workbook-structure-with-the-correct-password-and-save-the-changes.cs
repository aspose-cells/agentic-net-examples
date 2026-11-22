// Title: Remove workbook structure protection with a password using Aspose.Cells for .NET and save the unprotected file
// Common Searches: c# asp.net unprotect workbook structure with password using aspose.cells | how to load a password‑protected Excel file and save it unprotected with Aspose.Cells | example code for Workbook.Unprotect method in Aspose.Cells .NET | remove Excel workbook windows and structure protection programmatically in C# | asp.net core unprotect protected.xlsx and create unprotected.xlsx using Aspose.Cells
// Tags: Workbook.Unprotect password Aspose.Cells | unprotect workbook structure Aspose.Cells | remove Excel protection C# Aspose.Cells | save unprotected workbook Aspose.Cells | load password protected Excel file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads a password‑protected Excel workbook (protected.xlsx) with Aspose.Cells, calls Workbook.Unprotect using the supplied password to clear structure and window protection, then saves the resulting unprotected workbook as unprotected.xlsx, including file‑existence checks and exception handling.
class Program
{
    static void Main()
    {
        const string inputPath = "protected.xlsx";
        const string outputPath = "unprotected.xlsx";
        const string password = "YourPasswordHere";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found.");
            return;
        }

        try
        {
            // Load the protected workbook
            var workbook = new Workbook(inputPath);

            // Unprotect the workbook (structure and windows) using the password
            workbook.Unprotect(password);

            // Save the unprotected workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
