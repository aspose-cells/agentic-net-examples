using System;
using Aspose.Cells;

class VerifyFormulaUnchangedValueZero
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        Cell cell = ws.Cells["A1"];

        // Set a formula with its normal calculated result
        cell.SetFormula("=SUM(1,2)", 3);

        // Override the calculated result to zero while keeping the formula unchanged
        cell.SetFormula(cell.Formula, 0);

        // Verify that the formula has not changed
        Console.WriteLine("Formula: " + cell.Formula); // Expected: =SUM(1,2)

        // Verify that the cell value is now zero
        Console.WriteLine("Value before calculation: " + cell.Value); // Expected: 0

        // Recalculate to see the formula's true result
        wb.CalculateFormula();
        Console.WriteLine("Value after calculation: " + cell.Value); // Expected: 3

        // Save the workbook (optional)
        wb.Save("FormulaCheck.xlsx");
    }
}