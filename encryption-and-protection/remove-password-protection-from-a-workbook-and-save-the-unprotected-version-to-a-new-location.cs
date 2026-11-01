// Title: How to remove password protection from an Excel workbook and save it unprotected using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells remove password from .xlsx and save as new file | Open a password‑protected Excel workbook with Aspose.Cells and export it without a password | Aspose.Cells LoadOptions Password example for unprotecting Excel files in .NET | Programmatically delete workbook protection using Aspose.Cells in C# | Save an unprotected copy of a protected Excel workbook using Aspose.Cells for .NET
// Tags: Aspose.Cells remove workbook password | LoadOptions password property C# | save workbook without encryption Aspose.Cells | unprotect Excel file programmatically .NET | Aspose.Cells workbook decryption example | C# load protected xlsx Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample checks that the source .xlsx file exists, loads it with the supplied password via LoadOptions, and then saves a new copy without specifying a password, effectively removing workbook protection. Exceptions are caught and reported.
class RemoveWorkbookPassword
{
    static void Main()
    {
        // Path to the password‑protected workbook
        string inputFile = @"C:\Path\To\ProtectedWorkbook.xlsx";

        // Password required to open the workbook
        string workbookPassword = "YourPassword";

        // Path where the unprotected workbook will be saved
        string outputFile = @"C:\Path\To\UnprotectedWorkbook.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file not found: {inputFile}");
                return;
            }

            // Load the workbook using the password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = workbookPassword
            };
            Workbook workbook = new Workbook(inputFile, loadOptions);

            // Save the workbook without a password
            workbook.Save(outputFile, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved without password to: {outputFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
