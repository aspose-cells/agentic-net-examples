using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate source data
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);
        sheet.Cells["A3"].PutValue(30);

        // Add a formula that depends on the source data
        sheet.Cells["B1"].Formula = "=SUM(A1:A3)";

        // Copy rows 0‑2 (A1:C3) to rows starting at index 4 (row 5 in Excel)
        // Rule: Cells.CopyRows(sourceCells, sourceRowIndex, destinationRowIndex, rowNumber)
        sheet.Cells.CopyRows(sheet.Cells, 0, 4, 3);

        // Recalculate all formulas in the workbook so that copied rows reflect correct results
        // Rule: Workbook.CalculateFormula()
        workbook.CalculateFormula();

        // Save the workbook (lifecycle rule: save)
        workbook.Save("CopyRowsAndRecalc.xlsx");
    }
}