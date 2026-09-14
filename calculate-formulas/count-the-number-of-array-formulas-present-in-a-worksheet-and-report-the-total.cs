// Title: How to count both legacy CSE and dynamic array formulas in an Excel worksheet with Aspose.Cells for C#
// AI Prompts: Write C# code using Aspose.Cells that scans every cell in a worksheet and returns the total number of cells where IsArrayFormula or IsDynamicArrayFormula is true. | Extend the example to also gather the addresses of all cells that contain an array formula while still providing the overall count.
// Common Searches: Aspose.Cells count array formulas in .NET workbook | C# enumerate cells with IsDynamicArrayFormula property | detect legacy CSE array formulas using Aspose.Cells | retrieve total number of array formulas from Excel file with Aspose.Cells for C#
// Tags: count array formulas Aspose.Cells C# | detect legacy CSE array formula Aspose.Cells | identify dynamic array formulas in workbook | iterate worksheet cells Aspose.Cells | array formula enumeration .xlsx C#

using System;
using Aspose.Cells;

// The example creates a workbook, adds a legacy CSE array formula and a dynamic array formula, iterates through all cells checking IsArrayFormula and IsDynamicArrayFormula, counts the matching cells, prints the total count, and saves the workbook.
class CountArrayFormulas
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();               // lifecycle: create
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // -----------------------------------------------------------------
        // Sample setup: add a few array formulas so the demo has something
        // -----------------------------------------------------------------
        // Legacy (CSE) array formula that spills over three rows
        cells["A1"].SetArrayFormula("=SUM(B1:B3)", 3, 1);

        // Dynamic array formula (spills automatically)
        cells["C1"].SetDynamicArrayFormula("=SEQUENCE(2,2)", new FormulaParseOptions(), true);

        // ---------------------------------------------------------------
        // Count all cells that contain either a legacy or a dynamic array formula
        // ---------------------------------------------------------------
        int arrayFormulaCount = 0;
        foreach (Cell cell in cells)
        {
            if (cell.IsArrayFormula || cell.IsDynamicArrayFormula)
            {
                arrayFormulaCount++;
            }
        }

        Console.WriteLine("Total array formulas in the worksheet: " + arrayFormulaCount);

        // Save the workbook (optional, demonstrates lifecycle: save)
        workbook.Save("ArrayFormulaCountDemo.xlsx");
    }
}
