// Title: Remove workbook structure and windows protection from an Excel file using Aspose.Cells for .NET (C#) with a password
// Common Searches: C# Aspose.Cells example to unprotect workbook structure with password | How to remove Excel workbook windows protection using Aspose.Cells .NET | Aspose.Cells Workbook.Unprotect method usage for protected.xlsx | Make Excel sheets movable after removing protection with Aspose.Cells C# | Unprotect a password‑protected Excel workbook and save it unprotected using Aspose.Cells
// Tags: Aspose.Cells Workbook.Unprotect password | remove Excel workbook structure protection C# | unprotect workbook windows protection Aspose.Cells | save unprotected Excel file Aspose.Cells .NET | make sheets movable after unprotect Aspose.Cells | load protected.xlsx and save unprotected.xlsx C#

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for 'protected.xlsx', loads it with Aspose.Cells, calls Workbook.Unprotect with the supplied password to clear both structure and windows protection, and then saves the workbook as 'unprotected.xlsx', confirming that the worksheets are no longer locked and can be moved.
class Program
{
    static void Main()
    {
        const string inputPath = "protected.xlsx";
        const string outputPath = "unprotected.xlsx";
        const string password = "myPassword";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the protected workbook
            Workbook workbook = new Workbook(inputPath);

            // Remove workbook protection (structure, windows) using the correct API
            workbook.Unprotect(password);

            // Save the workbook without protection
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
