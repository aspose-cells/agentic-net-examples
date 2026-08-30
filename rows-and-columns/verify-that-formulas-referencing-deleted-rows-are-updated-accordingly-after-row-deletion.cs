// Title: Verify that a shared SUM formula updates its range automatically after deleting rows using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that creates a shared SUM formula, deletes a specific row, and prints the revised formula to demonstrate automatic range adjustment. | Show how to delete multiple rows with Aspose.Cells and retrieve the updated shared formula to confirm that the referenced range has shifted correctly.
// Common Searches: Aspose.Cells C# how does a shared formula adjust when a row is removed | example of deleting rows and keeping SUM formula correct with Aspose.Cells | C# code to verify formula range changes after DeleteRow in Aspose.Cells | shared formula reference update after DeleteRows Aspose.Cells .NET | testing formula auto‑update after row deletion using Aspose.Cells
// Tags: shared formula range adjustment Aspose.Cells | DeleteRow updates formula references .NET | Aspose.Cells automatic formula recalculation after row deletion | C# verify SUM formula shift after DeleteRows | Excel workbook formula integrity with Aspose.Cells

using System;
using Aspose.Cells;

// The sample creates a workbook, fills column B rows 10‑66 with incremental numbers, sets a shared SUM formula in A66 that references B10:B66, deletes row 9 and then rows 60‑61, and prints the formula after each deletion to show the range automatically updates. The workbook is saved for manual inspection.
class VerifyFormulaUpdateAfterRowDeletion
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate column B (index 1) with sample numeric values from row 10 to row 66
        // Row indices are zero‑based, so row 10 is index 9 and row 66 is index 65
        for (int i = 9; i <= 65; i++)
        {
            cells[i, 1].PutValue(i - 8); // simple incremental values
        }

        // Set a shared formula in row 66 (index 65) column A (index 0)
        // The formula sums the range B10:B66
        cells[65, 0].SetSharedFormula("=SUM(B10:B66)", 1, 6);

        // Display the original formula
        Console.WriteLine($"Original formula in A66: {cells[65, 0].Formula}");

        // -------------------------------------------------
        // Delete row 9 (index 8). This shifts all rows below up by one.
        // -------------------------------------------------
        cells.DeleteRow(8);

        // After deletion, the original row 66 becomes row 65 (index 64)
        Console.WriteLine($"After deleting row 9, formula in A65: {cells[64, 0].Formula}");

        // -------------------------------------------------
        // Delete rows 60‑61 (original rows). After the previous deletion,
        // those rows are now at indices 58 and 59.
        // -------------------------------------------------
        cells.DeleteRows(58, 2);

        // After this second deletion, the formula row moves up again:
        // original row 66 -> row 63 (index 62)
        Console.WriteLine($"After deleting rows 60‑61, formula in A63: {cells[62, 0].Formula}");

        // Save the workbook (optional, for manual inspection)
        workbook.Save("FormulaUpdateAfterDeletion.xlsx");
    }
}
