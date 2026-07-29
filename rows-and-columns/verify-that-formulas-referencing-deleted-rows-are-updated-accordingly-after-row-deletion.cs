// Title: C# – Verify automatic formula adjustment after row deletions with Aspose.Cells
// Description: Creates a workbook, fills column B with values 1‑10, sets a SUM formula in C11 (B1:B10), deletes row 5 and checks the formula updates to =SUM(B1:B9), then deletes rows 7‑8 and verifies the formula becomes =SUM(B1:B7), finally saves the file.
// Keywords: Aspose.Cells | C# | .NET | DeleteRow | DeleteRows | formula update | row deletion | SUM formula | cell reference adjustment | Excel automation
// Common Searches: Aspose.Cells update formula after deleting rows | C# delete row keep SUM formula correct | How to adjust formula references with DeleteRows in Aspose.Cells | Validate formula range after row removal .NET | Automatic formula recalculation Aspose.Cells
// Developer Intent: Confirm that formulas automatically reflect new ranges when rows are removed.
// Use Cases: Create a workbook, insert a SUM formula, delete a single row and verify the range shrinks accordingly. | Delete multiple consecutive rows and ensure the dependent formula updates to the new range. | Save the workbook to inspect the adjusted formula in Excel or programmatically.
// AI Prompts: Generate C# code using Aspose.Cells that deletes rows and asserts that all dependent formulas are updated correctly. | Write a unit test in C# for Aspose.Cells that checks the formula string after each row deletion operation. | Explain how Aspose.Cells recalculates formula references when rows are removed and how to retrieve the updated formula programmatically.

using System;
using Aspose.Cells;

namespace VerifyFormulaUpdateAfterRowDeletion
{
    // Creates a workbook, fills column B with values 1‑10, sets a SUM formula in C11 (B1:B10), deletes row 5 and checks the formula updates to =SUM(B1:B9), then deletes rows 7‑8 and verifies the formula becomes =SUM(B1:B7), finally saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data in column B (values 1..10)
            for (int i = 0; i < 10; i++)
            {
                cells[i, 1].PutValue(i + 1); // B1..B10
            }

            // Set a formula in cell C11 that sums B1:B10
            // This formula will reference rows that we will delete later
            cells[10, 2].Formula = "=SUM(B1:B10)"; // C11

            Console.WriteLine("Original formula in C11: " + cells[10, 2].Formula);

            // Delete row 5 (zero‑based index 4). Rows below shift up.
            // Use the DeleteRow method as defined in the documentation.
            cells.DeleteRow(4);

            // After deletion, the original B1:B10 range becomes B1:B9
            // The formula should be automatically adjusted.
            Console.WriteLine("After deleting row 5, formula in C10 (shifted up): " + cells[9, 2].Formula);

            // Verify the adjusted formula string (optional assertion)
            if (cells[9, 2].Formula != "=SUM(B1:B9)")
                Console.WriteLine("Warning: Formula was not updated as expected after single row deletion.");

            // Delete rows 7‑8 (zero‑based indices 6 and 7) – two rows total
            // Use the DeleteRows overload that takes row index, total rows.
            cells.DeleteRows(6, 2);

            // After this deletion, the range referenced by the formula should become B1:B7
            // The cell containing the formula has also shifted up due to row deletions.
            Console.WriteLine("After deleting rows 7‑8, formula in C8 (shifted up): " + cells[7, 2].Formula);

            // Verify the final adjusted formula string
            if (cells[7, 2].Formula != "=SUM(B1:B7)")
                Console.WriteLine("Warning: Formula was not updated as expected after multiple rows deletion.");

            // Save the workbook to verify manually if needed
            workbook.Save("FormulaUpdateAfterRowDeletion.xlsx");
        }
    }
}
