// Title: Remove password protection from an Excel workbook and all its worksheets using Aspose.Cells for .NET (C#)
// Common Searches: C# code to unprotect a password‑protected Excel workbook with Aspose.Cells | How to remove workbook and worksheet protection in Aspose.Cells .NET | Programmatically delete Excel file password using Aspose.Cells API | Aspose.Cells unprotect all sheets and save as unprotected file | Remove workbook structure protection without knowing the password in C#
// Tags: Aspose.Cells workbook unprotect C# | remove Excel password Aspose.Cells .NET | unprotect all worksheets Aspose.Cells | disable workbook structure protection C# | save unprotected Excel file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads a password‑protected Excel workbook with Aspose.Cells, calls Workbook.Unprotect and Worksheet.Unprotect (using an empty password) to clear workbook and sheet protections, then saves the file as an unprotected .xlsx.
class RemoveWorkbookProtection
{
    static void Main()
    {
        // Path to the protected workbook
        string inputPath = "protected.xlsx";
        // Path where the unprotected workbook will be saved
        string outputPath = "unprotected.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Remove workbook-level protection (structure, windows)
            // If the workbook was protected with a password, provide it here.
            // Using an empty string attempts to remove protection when no password is set.
            workbook.Unprotect("");

            // Iterate through all worksheets and remove their protection
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Remove worksheet protection; empty password works for unprotected sheets
                sheet.Unprotect("");
            }

            // Save the workbook without any protection
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
