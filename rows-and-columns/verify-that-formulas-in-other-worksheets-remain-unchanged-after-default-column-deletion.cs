using System;
using Aspose.Cells;

class VerifyFormulaUnchanged
{
    static void Main()
    {
        // Create a new workbook with two worksheets
        Workbook workbook = new Workbook();
        Worksheet sheet1 = workbook.Worksheets[0];               // Sheet1
        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");   // Sheet2

        // Fill some data in Sheet1 column A
        sheet1.Cells["A1"].PutValue(10);
        sheet1.Cells["A2"].PutValue(20);

        // In Sheet2 set a formula that references Sheet1!A1:A2
        sheet2.Cells["B1"].Formula = "=SUM(Sheet1!A1:A2)";

        // Calculate formulas so that dependent values are evaluated
        workbook.CalculateFormula();

        // Keep the original formula text for later comparison
        string originalFormula = sheet2.Cells["B1"].Formula;

        // Delete column A (index 0) from Sheet1 without updating references in other worksheets
        sheet1.Cells.DeleteColumn(0, false);

        // Re‑calculate formulas after the column deletion
        workbook.CalculateFormula();

        // Output the formula before and after deletion and verify it has not changed
        Console.WriteLine("Original formula : " + originalFormula);
        Console.WriteLine("Formula after deletion : " + sheet2.Cells["B1"].Formula);
        Console.WriteLine("Formula unchanged : " + (originalFormula == sheet2.Cells["B1"].Formula));

        // Save the workbook (optional, just to complete the lifecycle)
        workbook.Save("FormulaCheck.xlsx");
    }
}