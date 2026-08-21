// Title: C# – Recalculate all formulas after deleting rows and columns with Aspose.Cells
// Description: Creates a workbook, adds values and inter‑dependent formulas, deletes column B and the first row with UpdateReference enabled, then runs workbook.CalculateFormula() to update every formula and prints the results.
// Keywords: Aspose.Cells C# | CalculateFormula | DeleteColumns UpdateReference | DeleteRows UpdateReference | recalculate formulas after deletion | .NET spreadsheet manipulation | adjust cell references after column removal | adjust cell references after row removal
// Common Searches: Aspose.Cells recalculate formulas after deleting rows | UpdateReference effect on formulas Aspose.Cells .NET | How to delete a column and recalc formulas in C# | Calculate all formulas after structural changes Aspose.Cells
// Developer Intent: Update every formula in a workbook after removing a column and a row.
// Use Cases: Maintain correct calculation results when cleaning up a worksheet by deleting rows or columns. | Automate bulk removal of data while preserving dependent formula integrity before exporting. | Programmatically adjust spreadsheet layout and instantly obtain updated cell values.
// AI Prompts: Write C# code that deletes a column and a row with DeleteOptions.UpdateReference=true, then calls workbook.CalculateFormula() using Aspose.Cells. | Explain how DeleteOptions.UpdateReference updates cell references when rows or columns are removed in Aspose.Cells. | Show how to read and display calculated cell values after deletions and formula recalculation with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaAfterDeletion
{
    // Creates a workbook, adds values and inter‑dependent formulas, deletes column B and the first row with UpdateReference enabled, then runs workbook.CalculateFormula() to update every formula and prints the results.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            cells["A1"].PutValue(10);   // Value used in formulas
            cells["B1"].PutValue(20);
            cells["C1"].Formula = "=A1+B1";   // Formula 1
            cells["A2"].PutValue(5);
            cells["B2"].Formula = "=A2*2";    // Formula 2
            cells["C2"].Formula = "=C1+B2";   // Formula 3 (depends on previous formulas)

            // First deletion: delete column B (index 1)
            DeleteOptions delOpts = new DeleteOptions { UpdateReference = true };
            cells.DeleteColumns(1, 1, delOpts); // Removes column B and updates formulas

            // Second deletion: delete the first row (index 0)
            cells.DeleteRows(0, 1, delOpts); // Removes row 1 and updates formulas

            // After the two deletions, calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Display the resulting values after calculation
            Console.WriteLine("C1 value after deletions and calculation: " + cells["C1"].Value);
            Console.WriteLine("B2 value after deletions and calculation: " + cells["B2"].Value);
            Console.WriteLine("C2 value after deletions and calculation: " + cells["C2"].Value);
        }
    }
}
