// Title: Recalculate all formulas after deleting a column and a row with UpdateReference in Aspose.Cells for .NET
// Description: C# example that creates a workbook, adds values and formulas, deletes column B and the first row using DeleteOptions.UpdateReference, recalculates every formula with Workbook.CalculateFormula, prints the new results for C1 and C2, and saves the file as Result.xlsx.
// Keywords: Aspose.Cells calculate formulas | DeleteOptions.UpdateReference | delete column Aspose.Cells C# | delete row Aspose.Cells C# | recalculate workbook after structural changes | Aspose.Cells .NET formula update | Workbook.CalculateFormula example
// Common Searches: Aspose.Cells recalculate formulas after column deletion | UpdateReference option delete rows Aspose.Cells .NET | How to recalculate all formulas after removing rows and columns in Aspose.Cells | C# Aspose.Cells delete column and recalculate formulas
// Developer Intent: Recalculate every formula in a workbook after removing a column and a row while automatically updating cell references.
// Use Cases: Adjust formulas when cleaning up a worksheet by deleting unwanted columns or rows. | Generate a final report that requires structural changes before computing accurate results. | Automate Excel processing where rows/columns are removed and all dependent calculations must be refreshed.
// AI Prompts: Provide C# code using Aspose.Cells that deletes column B and the first row with DeleteOptions.UpdateReference=true, then recalculates all formulas. | Explain how DeleteOptions.UpdateReference affects cell references when rows or columns are removed in Aspose.Cells. | Show how to read and display the recalculated values of cells C1 and C2 after deletions in an Aspose.Cells workbook.

using System;
using Aspose.Cells;

// C# example that creates a workbook, adds values and formulas, deletes column B and the first row using DeleteOptions.UpdateReference, recalculates every formula with Workbook.CalculateFormula, prints the new results for C1 and C2, and saves the file as Result.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet and its cells
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data and formulas
        cells["A1"].PutValue(10);
        cells["B1"].PutValue(20);
        cells["C1"].Formula = "=A1+B1"; // Expected 30

        cells["A2"].PutValue(5);
        cells["B2"].PutValue(15);
        cells["C2"].Formula = "=A2*B2"; // Expected 75

        // First deletion: delete column B (index 1)
        DeleteOptions deleteOpts1 = new DeleteOptions();
        deleteOpts1.UpdateReference = true; // update formulas that reference the deleted column
        cells.DeleteColumns(1, 1, deleteOpts1);

        // Second deletion: delete the first row (index 0)
        DeleteOptions deleteOpts2 = new DeleteOptions();
        deleteOpts2.UpdateReference = true; // update formulas that reference the deleted row
        cells.DeleteRows(0, 1, deleteOpts2);

        // Calculate all formulas after the second deletion (rule: CalculateFormula)
        workbook.CalculateFormula();

        // Display the calculated results
        Console.WriteLine("C1 value after deletions and calculation: " + cells["C1"].Value);
        Console.WriteLine("C2 value after deletions and calculation: " + cells["C2"].Value);

        // Save the workbook (lifecycle rule: save)
        workbook.Save("Result.xlsx");
    }
}
