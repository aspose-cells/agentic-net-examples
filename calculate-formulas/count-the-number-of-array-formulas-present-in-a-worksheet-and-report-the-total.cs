// Title: Count Legacy and Dynamic Array Formulas in Aspose.Cells (C#)
// Description: Creates a workbook, adds sample data, inserts a CSE array formula with SetArrayFormula and a dynamic array with SetDynamicArrayFormula, calculates the sheet, scans all cells using IsArrayFormula and IsDynamicArrayFormula, outputs the total count, and saves the file.
// Keywords: Aspose.Cells | C# | array formula count | IsArrayFormula | IsDynamicArrayFormula | legacy array formula | dynamic array formula | SetArrayFormula | SetDynamicArrayFormula | .NET Excel automation
// Common Searches: count array formulas Aspose.Cells C# | detect legacy array formula Aspose.Cells | how to find dynamic array formulas in .NET | Aspose.Cells IsArrayFormula example | C# enumerate cells with array formulas
// Developer Intent: Obtain the total number of cells that contain any array formula in a worksheet.
// Use Cases: Verify that generated reports contain the expected number of array‑formula cells. | Audit workbooks to report array‑formula usage across multiple sheets. | Trigger custom logic (e.g., removal or replacement) based on the count of array formulas.
// AI Prompts: Generate C# code using Aspose.Cells that iterates through a worksheet and returns the count of cells where IsArrayFormula or IsDynamicArrayFormula is true. | Show how to extend the example to report separate counts for legacy and dynamic array formulas on each sheet of a multi‑sheet workbook. | Provide a snippet that logs the addresses of all array‑formula cells while still displaying the overall total.

using System;
using Aspose.Cells;

// Creates a workbook, adds sample data, inserts a CSE array formula with SetArrayFormula and a dynamic array with SetDynamicArrayFormula, calculates the sheet, scans all cells using IsArrayFormula and IsDynamicArrayFormula, outputs the total count, and saves the file.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle create)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate some sample data
        cells["A1"].PutValue(1);
        cells["A2"].PutValue(2);
        cells["A3"].PutValue(3);
        cells["B1"].PutValue(4);
        cells["B2"].PutValue(5);
        cells["B3"].PutValue(6);

        // Add a legacy (CSE) array formula that spans 3 rows, 1 column
        cells["B2"].SetArrayFormula("=A1:A3*2", 3, 1);

        // Add a dynamic array formula that will spill into 4 rows
        cells["C1"].SetDynamicArrayFormula("=SEQUENCE(4)", new FormulaParseOptions(), true);

        // Calculate formulas so that array results are materialized
        workbook.CalculateFormula();

        // Count cells that contain either a legacy array formula or a dynamic array formula
        int arrayFormulaCount = 0;
        foreach (Cell cell in cells)
        {
            if (cell.IsArrayFormula || cell.IsDynamicArrayFormula)
            {
                arrayFormulaCount++;
            }
        }

        Console.WriteLine("Total array formulas in the worksheet: " + arrayFormulaCount);

        // Save the workbook (lifecycle save)
        workbook.Save("ArrayFormulaCountDemo.xlsx");
    }
}
