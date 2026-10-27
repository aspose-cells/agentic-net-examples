// Title: How to remove workbook structure protection with a password and verify sheet addition using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# unprotect workbook structure with password | remove Excel workbook structure protection and add worksheet using Aspose.Cells | verify sheet addition after unprotecting a protected Excel file in .NET | C# code to load password‑protected .xlsx and disable structure protection with Aspose.Cells | unprotect Excel file and check if new sheets can be inserted using Aspose.Cells API
// Tags: Workbook.Unprotect password Aspose.Cells | remove workbook structure protection .xlsx | add worksheet after unprotecting Excel file | validate sheet count Aspose.Cells | handle protected Excel workbook C#

using Aspose.Cells;
using System;
using System.IO;

// The example loads a password‑protected Excel workbook (protected.xlsx), calls Workbook.Unprotect with the supplied password to clear structure and windows protection, adds a new worksheet to confirm that modifications are now allowed, and saves the result as unprotected.xlsx while outputting the sheet counts and handling any exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "protected.xlsx";
        const string outputPath = "unprotected.xlsx";
        const string password = "YourPasswordHere";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the protected workbook
            Workbook workbook = new Workbook(inputPath);

            // Unprotect the workbook using the password
            // This removes structure and windows protection
            workbook.Unprotect(password);

            // Verify that sheet modifications are now allowed
            int sheetCountBefore = workbook.Worksheets.Count;
            workbook.Worksheets.Add("NewSheet");
            int sheetCountAfter = workbook.Worksheets.Count;

            Console.WriteLine($"Sheets before adding: {sheetCountBefore}, after adding: {sheetCountAfter}");

            // Save the workbook with the structure unprotected
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
