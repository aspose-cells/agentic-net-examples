// Title: Recalculate all formulas in an Aspose.Cells workbook after deleting a column and a row with UpdateReference enabled (C#)
// AI Prompts: Delete column B and row 1 using DeleteOptions.UpdateReference, then call Workbook.CalculateFormula to refresh every dependent formula in Aspose.Cells (C#). | Show how to preserve formula references when removing rows or columns and subsequently recalculate the entire sheet with Aspose.Cells for .NET. | Demonstrate updating formula references after structural deletions and performing a full workbook calculation in C# using Aspose.Cells.
// Common Searches: Aspose.Cells calculate all formulas after deleting a column with UpdateReference in C# | How to keep formula references when removing rows in an Aspose.Cells workbook | C# example for DeleteRows and DeleteColumns with reference update and workbook recalculation | Recalculate dependent formulas after structural changes in Aspose.Cells .NET | Update formula links after column deletion and then recalc workbook using Aspose.Cells
// Tags: DeleteOptions.UpdateReference formula preservation Aspose.Cells | calculate workbook formulas after column deletion C# | recalculate after row removal Aspose.Cells | Aspose.Cells structural changes formula recalculation | C# DeleteRows DeleteColumns with reference update example

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaAfterDeletion
{
    // The example creates a workbook, adds values and formulas, deletes column B and then row 1 with UpdateReference enabled, recalculates all formulas, prints the updated cell values, and saves the workbook as an XLSX file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data and formulas
            cells["A1"].PutValue(10);          // Value
            cells["B1"].PutValue(20);          // Value
            cells["C1"].Formula = "=A1+B1";    // Formula dependent on A1 and B1
            cells["A2"].PutValue(5);
            cells["B2"].Formula = "=A2*2";     // Formula dependent on A2
            cells["C2"].Formula = "=C1+B2";    // Formula dependent on previous formulas

            // First deletion: delete column B (index 1)
            // Use DeleteOptions to update references in formulas
            DeleteOptions delOpts = new DeleteOptions { UpdateReference = true };
            cells.DeleteColumns(1, 1, delOpts); // Deletes column B

            // Second deletion: delete row 1 (index 0)
            // Again update references so formulas stay consistent
            cells.DeleteRows(0, 1, delOpts); // Deletes first row

            // After the second deletion, calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Output the results to verify calculation
            Console.WriteLine("After deletions and calculation:");
            Console.WriteLine($"C1 value (now at C0 after row deletion): {cells["C1"].Value}");
            Console.WriteLine($"B2 value (now at B1 after row deletion): {cells["B1"].Value}");
            Console.WriteLine($"C2 value (now at C1 after row deletion): {cells["C1"].Value}");

            // Save the workbook (lifecycle: save)
            workbook.Save("ResultAfterDeletion.xlsx");
        }
    }
}
