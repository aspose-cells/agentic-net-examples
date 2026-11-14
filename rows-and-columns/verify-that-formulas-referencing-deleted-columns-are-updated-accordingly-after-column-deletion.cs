// Title: Check that Aspose.Cells for .NET automatically updates SUM formulas when a worksheet column is deleted
// Common Searches: Aspose.Cells C# update formula range after deleting column B | how to ensure Excel formulas adjust after column removal using Aspose.Cells | C# verify SUM(A1:C1) changes to SUM(A1:B1) after DeleteColumn | recalculate workbook formulas after column deletion Aspose.Cells .NET | detect formula reference changes post column delete in Aspose.Cells
// Tags: DeleteColumn method formula reference update | Aspose.Cells automatic formula adjustment | SUM range correction after column removal | C# workbook recalculate formulas after column delete | Excel column deletion impact on formulas Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a workbook, fills cells A1‑C1, sets D1 to =SUM(A1:C1), calculates the result, deletes column B with DeleteColumn, recalculates, and confirms that the formula in D1 is automatically changed to =SUM(A1:B1) with the correct new value, then saves the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate columns A, B, C with sample values
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["B1"].PutValue(20);
        sheet.Cells["C1"].PutValue(30);

        // Set a formula in D1 that references the range A1:C1 (includes column B)
        sheet.Cells["D1"].Formula = "=SUM(A1:C1)";

        // Calculate formulas before any deletion
        workbook.CalculateFormula();
        Console.WriteLine("Before deletion - D1 value: " + sheet.Cells["D1"].Value); // Expected 60

        // Delete column B (zero‑based index 1)
        sheet.Cells.DeleteColumn(1);

        // Recalculate formulas after column deletion
        workbook.CalculateFormula();

        // Verify that the formula has been updated automatically
        Console.WriteLine("After deletion - D1 formula: " + sheet.Cells["D1"].Formula); // Expected "=SUM(A1:B1)"
        Console.WriteLine("After deletion - D1 value: " + sheet.Cells["D1"].Value);   // Expected 40 (10 + 30)

        // Save the workbook (optional, just to demonstrate lifecycle usage)
        workbook.Save("UpdatedFormulas.xlsx");
    }
}
