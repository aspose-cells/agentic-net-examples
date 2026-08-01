// Title: Automatic formula adjustment after row deletion with Aspose.Cells for .NET (C#)
// Description: This C# example creates a workbook, populates columns A and B, inserts a SUM formula in C10 that references B2:B9, then deletes rows using DeleteRow and DeleteRows with the updateReference flag set to true. After each deletion the formula range is automatically revised (B2:B8, then B2:B6). The workbook is saved for visual verification.
// Keywords: Aspose.Cells | C# | .NET | DeleteRow | DeleteRows | updateReference | formula adjustment | row deletion | Excel formula update | workbook manipulation
// Common Searches: Aspose.Cells update formula after deleting rows C# | DeleteRow keep formula references Aspose.Cells | DeleteRows updateReference example | How to adjust Excel formulas when rows are removed using Aspose.Cells | C# code to verify formula changes after row deletion
// Developer Intent: Confirm that Aspose.Cells automatically rewrites cell formulas to reflect the new range when rows are removed.
// Use Cases: Check that a formula '=SUM(B2:B9)' becomes '=SUM(B2:B8)' after deleting row 5 with DeleteRow(updateReference:true). | Validate that the same formula updates to '=SUM(B2:B6)' after removing rows 7‑8 via DeleteRows(updateReference:true). | Save the workbook and open it in Excel to see the corrected formula references.
// AI Prompts: Generate a C# unit test using Aspose.Cells that asserts the formula in C10 changes from '=SUM(B2:B9)' to '=SUM(B2:B8)' after DeleteRow(updateReference:true). | Write code that logs the formula before and after deleting rows 7‑8 with DeleteRows(updateReference:true) and verifies the expected range. | Explain how the DeleteRow and DeleteRows methods with the updateReference parameter modify existing formulas in an Aspose.Cells workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaUpdateDemo
{
    // This C# example creates a workbook, populates columns A and B, inserts a SUM formula in C10 that references B2:B9, then deletes rows using DeleteRow and DeleteRows with the updateReference flag set to true. After each deletion the formula range is automatically revised (B2:B8, then B2:B6). The workbook is saved for visual verification.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data in column A (values) and column B (values)
            for (int i = 0; i < 10; i++)
            {
                cells[i, 0].PutValue(i + 1);          // A1:A10 = 1..10
                cells[i, 1].PutValue((i + 1) * 10);   // B1:B10 = 10,20,...,100
            }

            // Set a formula in row 10 that sums B2:B9 (i.e., rows 2 through 9)
            // This formula will be affected when rows are deleted.
            cells[9, 2].Formula = "=SUM(B2:B9)"; // C10

            Console.WriteLine("Before deletion:");
            Console.WriteLine($"C10 formula: {cells[9, 2].Formula}");

            // Delete a single row (row index 4 corresponds to Excel row 5)
            // Use the overload that updates references automatically.
            cells.DeleteRow(4, true);

            // After deleting row 5, the original range B2:B9 becomes B2:B8.
            // The formula in C10 should be adjusted accordingly.
            Console.WriteLine("\nAfter deleting row 5:");
            Console.WriteLine($"C10 formula: {cells[8, 2].Formula}"); // C9 now holds the original C10 cell

            // Delete multiple rows: delete rows 6 and 7 (original Excel rows 7 and 8)
            // Use DeleteRows with updateReference = true.
            cells.DeleteRows(5, 2, true);

            // After the second deletion, the formula should now reference B2:B6.
            Console.WriteLine("\nAfter deleting rows 7-8:");
            Console.WriteLine($"C8 formula: {cells[7, 2].Formula}"); // C8 now holds the formula cell

            // Save the workbook to verify the changes visually if needed.
            workbook.Save("FormulaUpdateAfterRowDeletion.xlsx");
        }
    }
}
