// Title: Recalculate all formulas in an Aspose.Cells workbook after deleting a row using C#
// AI Prompts: Delete a specific row in a worksheet and automatically update all dependent formulas with Aspose.Cells for .NET. | Refresh the entire workbook’s calculations after removing rows to reflect changed cell references in C#.
// Common Searches: Aspose.Cells recalculate formulas after row deletion in C# | How to update SUM range when a row is removed using Aspose.Cells .NET | C# delete a row and refresh Excel formulas with Aspose.Cells library | Programmatically recalculate workbook after modifying rows Aspose.Cells | Calculate all formulas after deleting rows in an Excel file using Aspose.Cells
// Tags: Aspose.Cells delete row with formula update | C# recalculate workbook formulas | update Excel formula references after row removal | Aspose.Cells CalculateFormula after data change | Excel SUM range adjustment using Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a workbook, fills cells A1‑A3, sets B1 to =SUM(A1:A3), calculates the formula, deletes the second row while updating references, recalculates all formulas, prints the before and after values of B1, and saves the result as DeletedRowCalc.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Fill some data
        cells["A1"].PutValue(10);
        cells["A2"].PutValue(20);
        cells["A3"].PutValue(30);

        // Add a formula that sums the three values
        cells["B1"].Formula = "=SUM(A1:A3)";

        // Initial calculation (optional, shows original result)
        workbook.CalculateFormula();
        Console.WriteLine("Before deletion, B1 = " + cells["B1"].Value);

        // Delete the second row (index 1). Update references so the formula adjusts.
        cells.DeleteRow(1, true);

        // Re‑calculate all formulas after the deletion operation
        workbook.CalculateFormula();
        Console.WriteLine("After deletion, B1 = " + cells["B1"].Value);

        // Save the workbook
        workbook.Save("DeletedRowCalc.xlsx");
    }
}
