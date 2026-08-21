// Title: Aspose.Cells for .NET – Verify Shared SUM Formula Updates After Row Deletion
// Description: C# example that creates a workbook, assigns a shared formula "=SUM(B10:B66)" to A66, deletes row 9 and rows 60‑61, then prints the formulas to show the range automatically changes to B10:B65 and B10:B63 before saving the file.
// Keywords: Aspose.Cells .NET | C# shared formula | row deletion formula update | adjust formula range | SUM formula after delete rows | Aspose.Cells example | GitHub Aspose.Cells | Excel automation C#
// Common Searches: Aspose.Cells update formula after deleting rows | C# shared formula adjusts when rows are removed | How to keep SUM range correct after row deletion in Aspose.Cells | Aspose.Cells example for automatic formula shift | GitHub sample Aspose.Cells row deletion
// Developer Intent: Confirm that Aspose.Cells automatically rewrites shared formulas to reflect new cell ranges when rows are deleted.
// Use Cases: Validate that a shared SUM formula changes from B10:B66 to B10:B65 after removing row 9. | Ensure the same formula further updates to B10:B63 after deleting rows 60‑61. | Display the revised formula strings from the affected cells to demonstrate correct adjustment.
// AI Prompts: Generate a unit test in C# using Aspose.Cells that asserts the formula in A66 becomes "=SUM(B10:B65)" after deleting row 9. | Show C# code that compares the original and updated formula texts after deleting rows 60‑61 with Aspose.Cells. | Explain the internal mechanism Aspose.Cells uses to recalculate shared formula references when rows are removed and how to retrieve the updated formula.

using System;
using Aspose.Cells;

namespace FormulaUpdateAfterRowDeletion
{
    // C# example that creates a workbook, assigns a shared formula "=SUM(B10:B66)" to A66, deletes row 9 and rows 60‑61, then prints the formulas to show the range automatically changes to B10:B65 and B10:B63 before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();
            Cells cells = workbook.Worksheets[0].Cells;

            // ------------------------------------------------------------
            // 1. Set a shared formula that references a range including row 66
            // ------------------------------------------------------------
            // Row index is zero‑based, so row 66 is index 65.
            // The formula sums B10:B66 (B column, rows 10‑66).
            cells[65, 0].SetSharedFormula("=SUM(B10:B66)", 1, 6); // A66

            // ------------------------------------------------------------
            // 2. Delete a single row (row 9, zero‑based index 8)
            //    This should shift rows up and adjust the formula range to B10:B65
            // ------------------------------------------------------------
            cells.DeleteRow(8); // Delete row 9

            // Verify the adjusted formula after the first deletion
            Console.WriteLine("After deleting row 9:");
            for (int col = 1; col <= 6; col++) // Columns B‑G (indexes 1‑6)
            {
                string cellName = CellsHelper.CellIndexToName(64, col); // Row 65 after shift
                Console.WriteLine($"{cellName} formula: {cells[64, col].Formula}");
            }

            // ------------------------------------------------------------
            // 3. Delete two rows (rows 60‑61, zero‑based indexes 59‑60)
            //    After the previous deletion the original row 66 became row 65.
            //    Deleting rows 60‑61 will shift the formula range to B10:B63.
            // ------------------------------------------------------------
            cells.DeleteRows(59, 2); // Delete rows 60 and 61

            // Verify the adjusted formula after the second deletion
            Console.WriteLine("\nAfter deleting rows 60‑61:");
            for (int col = 1; col <= 6; col++) // Columns B‑G
            {
                string cellName = CellsHelper.CellIndexToName(62, col); // Row 63 after shift
                Console.WriteLine($"{cellName} formula: {cells[62, col].Formula}");
            }

            // ------------------------------------------------------------
            // 4. Save the workbook (optional, just to demonstrate lifecycle)
            // ------------------------------------------------------------
            workbook.Save("FormulaUpdateAfterRowDeletion.xlsx");
        }
    }
}
