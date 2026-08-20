// Title: Unprotect a worksheet, retain formatting rights, and update a cell formula with Aspose.Cells (C#)
// Description: Demonstrates how to read a protected worksheet, capture the AllowFormattingCell flag, unprotect it, modify the formula in a specific cell, restore the original formatting permission, recalculate formulas, and save the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells unprotect worksheet C# | preserve AllowFormattingCell permission | update cell formula Aspose.Cells | re‑protect worksheet with password | .NET spreadsheet protection | calculate formulas Aspose.Cells
// Common Searches: how to unprotect an Aspose.Cells worksheet without losing formatting rights | change a formula in a protected Excel file using Aspose.Cells C# | restore worksheet protection settings after editing formulas | Aspose.Cells keep AllowFormattingCell flag when updating cells
// Developer Intent: Remove protection, keep formatting permission, modify a formula, and re‑apply protection.
// Use Cases: Adjust calculations in a locked financial model while preserving user formatting access. | Batch‑process template workbooks that are password‑protected, updating formulas without resetting permissions. | Refresh report formulas in a secured workbook after data import, then re‑secure the sheet.
// AI Prompts: Generate C# code that unprotects an Aspose.Cells worksheet, saves the AllowFormattingCell setting, updates a given cell formula, and protects the sheet again with the same password. | Explain how to retrieve and restore worksheet protection options such as AllowFormattingCell when editing formulas with Aspose.Cells. | Create a reusable method that accepts a file path, password, cell address, and new formula, then performs unprotect‑update‑protect while preserving all original protection flags.

using System;
using Aspose.Cells;

// Demonstrates how to read a protected worksheet, capture the AllowFormattingCell flag, unprotect it, modify the formula in a specific cell, restore the original formatting permission, recalculate formulas, and save the workbook using Aspose.Cells for .NET.
class UnprotectAndUpdateFormula
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["A2"].PutValue(10);

            // Set an initial formula in B1
            sheet.Cells["B1"].Formula = "=SUM(A1:A2)";

            // Protect the worksheet with a password and allow cell formatting
            string password = "pwd123";
            sheet.Protect(ProtectionType.All, password, null);
            sheet.Protection.AllowFormattingCell = true; // preserve formatting permission

            // Preserve formatting permission before unprotecting
            bool allowFormattingCell = sheet.Protection.AllowFormattingCell;

            // Unprotect the worksheet using the password
            sheet.Unprotect(password);

            // Update the formula in B1
            sheet.Cells["B1"].Formula = "=A1*2";

            // Re‑protect the worksheet, restoring the formatting permission
            sheet.Protect(ProtectionType.All, password, null);
            sheet.Protection.AllowFormattingCell = allowFormattingCell;

            // Calculate formulas so the workbook stores the result values
            workbook.CalculateFormula();

            // Save the modified workbook
            workbook.Save("UnprotectedUpdated.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
