// Title: C# – Preserve Formulas When Deleting Columns with Aspose.Cells DeleteOptions.UpdateReference = false
// Description: Demonstrates how to keep original cell references unchanged after deleting a column in an Aspose.Cells workbook by setting DeleteOptions.UpdateReference to false. The example creates a workbook, adds values and a formula, applies the DeleteOptions, deletes column A, and saves the file, resulting in a #REF! error where the formula’s source cell was removed.
// Keywords: Aspose.Cells DeleteOptions | UpdateReference false | C# preserve formulas | delete column without adjusting formulas | #REF! error Aspose.Cells | .NET spreadsheet manipulation
// Common Searches: Aspose.Cells delete column keep formula reference | DeleteOptions.UpdateReference false C# example | prevent formula shift when deleting rows Aspose.Cells | how to get #REF! after column deletion in Aspose.Cells
// Developer Intent: Show how to disable automatic formula reference updates when removing rows or columns using DeleteOptions.UpdateReference = false in a .NET workbook.
// Use Cases: Audit spreadsheets by intentionally generating #REF! errors after column removal. | Remove data blocks from financial models without breaking dependent calculations. | Batch‑delete multiple columns while preserving original references for later reconciliation.
// AI Prompts: Write C# code that deletes a column in an Aspose.Cells workbook without updating any formula references using DeleteOptions.UpdateReference = false. | Explain the impact of DeleteOptions.UpdateReference on cell formulas when a column is removed, and provide a short code snippet. | Show how to configure DeleteOptions to keep original formulas intact while deleting a range of rows in a .NET spreadsheet.

using System;
using Aspose.Cells;

// Demonstrates how to keep original cell references unchanged after deleting a column in an Aspose.Cells workbook by setting DeleteOptions.UpdateReference to false. The example creates a workbook, adds values and a formula, applies the DeleteOptions, deletes column A, and saves the file, resulting in a #REF! error where the formula’s source cell was removed.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some cells with values and a formula that references them
        worksheet.Cells["A1"].PutValue(10);
        worksheet.Cells["B1"].PutValue(20);
        worksheet.Cells["C1"].Formula = "=A1+B1";

        // Create DeleteOptions and set UpdateReference to false
        // This ensures that when a row/column is deleted, existing formulas are NOT adjusted.
        DeleteOptions deleteOptions = new DeleteOptions
        {
            UpdateReference = false
        };

        // Delete column A (index 0) using the DeleteOptions.
        // Because UpdateReference is false, the formula in C1 will retain its original reference to A1,
        // which will now point to a non‑existent cell (resulting in #REF! in Excel).
        worksheet.Cells.DeleteColumns(0, 1, deleteOptions);

        // Save the workbook to a file
        workbook.Save("DeleteOptionsFalseDemo.xlsx");
    }
}
