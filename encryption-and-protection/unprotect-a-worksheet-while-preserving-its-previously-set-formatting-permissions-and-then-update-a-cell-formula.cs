// Title: C# – Unprotect Aspose.Cells Worksheet, Keep Formatting Permissions, and Update a Cell Formula
// Description: Demonstrates how to read a worksheet's AllowFormattingCell/Column/Row flags, temporarily unprotect the sheet with a password, modify a formula, recalculate, and save the workbook while restoring the original formatting permissions.
// Keywords: Aspose.Cells unprotect worksheet C# | preserve formatting permissions Aspose.Cells | update cell formula after unprotect | worksheet protection AllowFormattingCell | C# Excel protection Aspose.Cells
// Common Searches: how to keep AllowFormattingCell when unprotecting an Aspose.Cells sheet | C# unprotect worksheet without losing formatting rights | change formula in protected Excel sheet using Aspose.Cells | restore worksheet protection options after editing a formula | Aspose.Cells protect and unprotect worksheet example
// Developer Intent: Temporarily remove worksheet protection, retain the original formatting allowances, modify a cell's formula, recalculate, and save the file.
// Use Cases: Automated report pipelines that need to adjust formulas in protected sheets while preserving user formatting capabilities. | Batch processing of workbooks where formulas are refreshed without permanently disabling sheet protection. | Enterprise applications that allow end‑users to format cells but restrict structural changes, requiring occasional programmatic formula updates.
// AI Prompts: Write C# code with Aspose.Cells to read formatting permission flags, unprotect a worksheet, set a new formula in B2, recalculate, and reapply the same permissions. | Explain the relationship between Aspose.Cells Protection object properties (AllowFormattingCell, AllowFormattingColumn, AllowFormattingRow) and the Unprotect method. | Provide a step‑by‑step tutorial for protecting a worksheet with specific formatting rights, safely unprotecting it to edit a formula, and restoring the original protection settings.

using System;
using Aspose.Cells;

namespace AsposeCellsWorksheetUnprotectAndFormula
{
    // Demonstrates how to read a worksheet's AllowFormattingCell/Column/Row flags, temporarily unprotect the sheet with a password, modify a formula, recalculate, and save the workbook while restoring the original formatting permissions.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ----- Set up initial protection with formatting permissions -----
                // Allow formatting of cells while the sheet is protected
                sheet.Protection.AllowFormattingCell = true;
                sheet.Protection.AllowFormattingColumn = true;
                sheet.Protection.AllowFormattingRow = true;

                // Protect the worksheet with a password
                string password = "SecretPwd";
                sheet.Protect(ProtectionType.All, password, null);

                // ----- Preserve formatting permissions before unprotecting -----
                bool allowFmtCell = sheet.Protection.AllowFormattingCell;
                bool allowFmtColumn = sheet.Protection.AllowFormattingColumn;
                bool allowFmtRow = sheet.Protection.AllowFormattingRow;

                // ----- Unprotect the worksheet -----
                sheet.Unprotect(password);

                // Restore formatting permissions after unprotecting
                sheet.Protection.AllowFormattingCell = allowFmtCell;
                sheet.Protection.AllowFormattingColumn = allowFmtColumn;
                sheet.Protection.AllowFormattingRow = allowFmtRow;

                // ----- Update a cell formula -----
                // Example: set formula in cell B2 to sum of A1:A5
                Cell targetCell = sheet.Cells["B2"];
                // Use the Formula property to assign the formula string
                targetCell.Formula = "=SUM(A1:A5)";

                // Optionally calculate formulas so the result is stored
                workbook.CalculateFormula();

                // ----- Save the workbook -----
                string outputPath = "UnprotectedAndFormulaUpdated.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
