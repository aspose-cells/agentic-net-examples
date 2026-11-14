// Title: C# example: automatically update Excel formula references after deleting rows with Aspose.Cells
// Common Searches: asp.net keep Excel formulas correct after deleting a row using Aspose.Cells | c# delete a worksheet row and shift formula references automatically | aspose.cells recalculate formulas after DeleteRows method | verify formula reference updates when a row is removed in a .NET workbook | example of adjusting cell references after row deletion with Aspose.Cells C#
// Tags: Aspose.Cells DeleteRows formula reference update | C# workbook.CalculateFormula after row removal | Excel formula adjustment on row deletion Aspose | automatic formula refresh with Aspose.Cells | row deletion impact on cell references .NET

using System;
using Aspose.Cells;

// The sample creates a workbook, populates cells A2‑B3, assigns the formula "=A2+B2" to C1, calculates the result, deletes the second row using DeleteRows, recalculates the workbook, and demonstrates that the formula reference in C1 is automatically adjusted and re‑evaluated, confirming correct formula updates after row deletion.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        sheet.Cells["A2"].PutValue(10);
        sheet.Cells["B2"].PutValue(20);
        sheet.Cells["A3"].PutValue(30);
        sheet.Cells["B3"].PutValue(40);

        // Set a formula in C1 that references the second row (A2+B2)
        sheet.Cells["C1"].Formula = "=A2+B2";

        // Calculate formulas to obtain initial value
        workbook.CalculateFormula();

        Console.WriteLine("Before row deletion:");
        Console.WriteLine($"C1 formula: {sheet.Cells["C1"].Formula}");
        Console.WriteLine($"C1 value: {sheet.Cells["C1"].Value}");

        // Delete the second row (index 1 because rows are zero‑based)
        sheet.Cells.DeleteRows(1, 1);

        // Recalculate formulas after deletion
        workbook.CalculateFormula();

        Console.WriteLine("After row deletion:");
        Console.WriteLine($"C1 formula: {sheet.Cells["C1"].Formula}");
        Console.WriteLine($"C1 value: {sheet.Cells["C1"].Value}");

        // Save the workbook (optional verification)
        workbook.Save("FormulaUpdateAfterRowDeletion.xlsx");
    }
}
