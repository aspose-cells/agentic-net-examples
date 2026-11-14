// Title: C# example to verify that changing an Aspose.Cells worksheet protection password preserves formulas and values
// Common Searches: Aspose.Cells C# check that worksheet data remains unchanged after updating protection password | How to validate Excel cell formulas are intact after changing worksheet password with Aspose.Cells | C# code to compare cell values before and after unprotect/protect using Aspose.Cells | Ensure worksheet protection password change does not affect formulas in Aspose.Cells .NET | Detect data alteration when re‑protecting a sheet with a new password in Aspose.Cells
// Tags: worksheet protect password change validation Aspose.Cells | preserve cell formulas after worksheet unprotect Aspose.Cells | compare cell values before after protect Aspose.Cells | Excel data integrity verification Aspose.Cells .NET | unprotect protect worksheet without data loss C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, records each cell's formula and value, protects the first worksheet with an old password, then unprotects and re‑protects it using a new password, and finally iterates through all cells to confirm that neither formulas nor values have changed, reporting the result and saving the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string oldPassword = "oldPassword";
        const string newPassword = "newPassword";

        // Verify input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Capture current formulas and values of all used cells
            var originalData = new Dictionary<string, (string Formula, object Value)>();
            foreach (Cell cell in worksheet.Cells)
            {
                originalData[cell.Name] = (cell.Formula, cell.Value);
            }

            // Ensure the sheet is protected with the old password (if not already)
            // The third parameter (oldPassword) is required by the API; pass null for initial protection
            worksheet.Protect(ProtectionType.All, oldPassword, null);

            // Change the protection password
            worksheet.Unprotect(oldPassword); // remove old password
            // Apply new password; oldPassword parameter not needed after unprotect, pass null
            worksheet.Protect(ProtectionType.All, newPassword, null);

            // Verify that formulas and values remain unchanged after password change
            bool allUnchanged = true;
            foreach (Cell cell in worksheet.Cells)
            {
                var key = cell.Name;
                var (origFormula, origValue) = originalData[key];

                bool formulaSame = string.Equals(origFormula, cell.Formula, StringComparison.Ordinal);
                bool valueSame = (origValue == null && cell.Value == null) ||
                                 (origValue != null && origValue.Equals(cell.Value));

                if (!formulaSame || !valueSame)
                {
                    allUnchanged = false;
                    Console.WriteLine($"Cell {key} changed. Formula: '{origFormula}' -> '{cell.Formula}', Value: '{origValue}' -> '{cell.Value}'");
                }
            }

            Console.WriteLine(allUnchanged
                ? "Validation passed: No formulas or values were altered by changing the protection password."
                : "Validation failed: Some formulas or values were altered.");

            // Save the workbook (optional)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
