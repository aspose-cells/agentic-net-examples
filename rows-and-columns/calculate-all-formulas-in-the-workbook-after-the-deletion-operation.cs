using System;
using Aspose.Cells;

namespace AsposeCellsFormulaRecalcAfterDeletion
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].PutValue(30);
            cells["A4"].PutValue(40);

            // Add formulas that depend on the data range A1:A4
            cells["B1"].Formula = "=SUM(A1:A4)";   // Should be 100
            cells["B2"].Formula = "=AVERAGE(A1:A4)"; // Should be 25

            // Calculate formulas before deletion (optional, just to show initial state)
            workbook.CalculateFormula();

            Console.WriteLine("Before deletion:");
            Console.WriteLine($"B1 = {cells["B1"].Value}"); // Expected 100
            Console.WriteLine($"B2 = {cells["B2"].Value}"); // Expected 25

            // Delete the third row (index 2, zero‑based). Update references so formulas adjust.
            cells.DeleteRow(2, true); // Row 3 (A3) is removed; range becomes A1:A3

            // Recalculate all formulas after the deletion
            workbook.CalculateFormula();

            Console.WriteLine("\nAfter deletion of row 3:");
            // The formulas now refer to the updated range A1:A3 (10,20,40)
            Console.WriteLine($"B1 = {cells["B1"].Value}"); // Expected 70
            Console.WriteLine($"B2 = {cells["B2"].Value}"); // Expected 23.333...

            // Save the workbook to verify the results in Excel
            workbook.Save("FormulaRecalcAfterDeletion.xlsx");
        }
    }
}