using System;
using Aspose.Cells;

class CountArrayFormulas
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // Sample data to demonstrate both legacy and dynamic
        // array formulas. In real usage, load your workbook
        // and skip this setup.
        // -------------------------------------------------
        worksheet.Cells["A1"].PutValue(1);
        worksheet.Cells["A2"].PutValue(2);
        worksheet.Cells["A3"].PutValue(3);

        // Legacy (CSE) array formula spanning 3 rows, 1 column
        worksheet.Cells["B1"].SetArrayFormula("=A1:A3*2", 3, 1);

        // Dynamic array formula that spills into a 2x2 range
        worksheet.Cells["C1"].SetDynamicArrayFormula("=SEQUENCE(2,2)", new FormulaParseOptions(), true);
        // -------------------------------------------------

        // Count all cells that contain either a legacy array formula
        // or a dynamic array formula.
        int arrayFormulaCount = 0;
        Cells cells = worksheet.Cells;

        foreach (Cell cell in cells)
        {
            if (cell.IsArrayFormula || cell.IsDynamicArrayFormula)
            {
                arrayFormulaCount++;
            }
        }

        Console.WriteLine("Total array formulas in the worksheet: " + arrayFormulaCount);

        // Save the workbook (optional, for verification)
        workbook.Save("ArrayFormulaCountDemo.xlsx");
    }
}