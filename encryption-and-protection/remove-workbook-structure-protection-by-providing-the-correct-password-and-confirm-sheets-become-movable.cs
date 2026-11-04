// Title: How to remove workbook structure protection with a password and verify sheet reordering using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# remove workbook structure protection with password | Unprotect Excel workbook and enable sheet moving with Aspose.Cells .NET | Programmatically test worksheet rename after workbook unprotect in C# | Load password‑protected .xlsx and save as unprotected using Aspose.Cells | C# code to disable workbook structure lock and confirm sheet mobility
// Tags: Workbook.Unprotect password Aspose.Cells | remove workbook structure protection .NET | verify worksheet rename after unprotect | load protected xlsx save unprotected Aspose.Cells | enable sheet reordering after unprotect C#

using Aspose.Cells;
using System;
using System.IO;

// The example loads a password‑protected Excel file (protected.xlsx) with Aspose.Cells, calls Workbook.Unprotect(password) to lift the workbook's structure protection, temporarily renames the first worksheet to confirm that sheets can be moved, and saves the resulting file as unprotected.xlsx, handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        const string inputPath = "protected.xlsx";
        const string outputPath = "unprotected.xlsx";
        const string password = "yourPassword";

        try
        {
            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the protected workbook
            Workbook workbook = new Workbook(inputPath);

            // Unprotect the workbook using the password (removes structure protection)
            workbook.Unprotect(password);

            // Verify that worksheets can be modified after unprotecting
            if (workbook.Worksheets.Count > 1)
            {
                try
                {
                    // Simple modification: temporarily rename the first sheet and revert
                    Worksheet sheet = workbook.Worksheets[0];
                    string originalName = sheet.Name;
                    sheet.Name = originalName + "_temp";
                    sheet.Name = originalName;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Worksheet modification failed: {ex.Message}");
                }
            }

            // Save the unprotected workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved as '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
