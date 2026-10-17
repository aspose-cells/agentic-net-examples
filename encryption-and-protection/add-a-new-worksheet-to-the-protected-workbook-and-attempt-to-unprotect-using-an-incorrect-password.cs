// Title: Add a new worksheet to a password‑protected Excel workbook and attempt unprotect with an incorrect password using Aspose.Cells for .NET
// Common Searches: how to add a sheet to a password protected Excel workbook with Aspose.Cells C# | Aspose.Cells unprotect workbook with wrong password handling in .NET | load protected xlsx using LoadOptions password then modify worksheet C# | add worksheet to protected workbook without removing protection Aspose.Cells | exception handling when unprotecting Excel file with incorrect password Aspose.Cells
// Tags: add worksheet to protected workbook Aspose.Cells | unprotect Excel workbook with incorrect password C# | load protected Excel file using LoadOptions password | Aspose.Cells workbook protection API | save modified protected workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a password‑protected workbook if missing, loads it with the correct password via LoadOptions, attempts to unprotect it using an incorrect password (which fails silently), adds a new worksheet named "NewSheet" with a value in A1, and saves the updated file, demonstrating protection handling and exception management in Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        const string inputPath = "protected.xlsx";
        const string outputPath = "output.xlsx";
        const string correctPassword = "correctPassword";
        const string wrongPassword = "incorrectPassword";

        try
        {
            // Ensure the input file exists; if not, create a protected workbook for demo purposes
            if (!File.Exists(inputPath))
            {
                var wb = new Workbook();
                wb.Worksheets[0].Name = "Sheet1";

                // Protect the workbook with a known password
                wb.Protect(ProtectionType.All, correctPassword);
                wb.Save(inputPath);
            }

            // Verify the input file still exists before loading
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file '{inputPath}' was not found.");

            // Load the protected workbook using the correct password
            var loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = correctPassword
            };
            var workbook = new Workbook(inputPath, loadOptions);

            // Attempt to unprotect with an incorrect password (will fail silently)
            workbook.Unprotect(wrongPassword);

            // Add a new worksheet
            Worksheet newSheet = workbook.Worksheets.Add("NewSheet");
            newSheet.Cells["A1"].PutValue("This is a new sheet.");

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (CellsException ex)
        {
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
