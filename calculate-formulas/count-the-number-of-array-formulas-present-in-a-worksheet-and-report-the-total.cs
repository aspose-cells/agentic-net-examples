// Title: Count Legacy and Dynamic Array Formulas in Aspose.Cells (C#)
// Description: Creates a workbook, adds sample data, inserts a CSE array formula with SetArrayFormula and a dynamic array with SetDynamicArrayFormula, calculates the sheet, scans all cells using IsArrayFormula and IsDynamicArrayFormula, outputs the total count, and saves the file.
// Keywords: Aspose.Cells | C# | array formula count | IsArrayFormula | IsDynamicArrayFormula | legacy array formula | dynamic array formula | SetArrayFormula | SetDynamicArrayFormula | .NET Excel automation
// Common Searches: count array formulas Aspose.Cells C# | detect legacy array formula Aspose.Cells | how to find dynamic array formulas in .NET | Aspose.Cells IsArrayFormula example | C# enumerate cells with array formulas
// Developer Intent: Obtain the total number of cells that contain any array formula in a worksheet.
// Use Cases: Verify that generated reports contain the expected number of array‑formula cells. | Audit workbooks to report array‑formula usage across multiple sheets. | Trigger custom logic (e.g., removal or replacement) based on the count of array formulas.
// AI Prompts: Generate C# code using Aspose.Cells that iterates through a worksheet and returns the count of cells where IsArrayFormula or IsDynamicArrayFormula is true. | Show how to extend the example to report separate counts for legacy and dynamic array formulas on each sheet of a multi‑sheet workbook. | Provide a snippet that logs the addresses of all array‑formula cells while still displaying the overall total.

using System;
using Aspose.Cells;

namespace AsposeCellsArrayFormulaCounter
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // ----- Sample data (optional) -----
            // Populate some values
            cells["A1"].PutValue(1);
            cells["A2"].PutValue(2);
            cells["A3"].PutValue(3);
            cells["B1"].PutValue(4);
            cells["B2"].PutValue(5);
            cells["B3"].PutValue(6);

            // Add a legacy (CSE) array formula in cell C1 that spills to C1:C3
            cells["C1"].SetArrayFormula("=A1:A3*2", 3, 1);

            // Add a dynamic array formula in cell D1 (will spill horizontally)
            cells["D1"].SetDynamicArrayFormula("=TRANSPOSE(A1:A3)", new FormulaParseOptions(), true);

            // Calculate formulas so that array results are materialized
            workbook.CalculateFormula();

            // ----- Count array formulas -----
            int arrayFormulaCount = 0;

            // Determine the used range of the worksheet
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Iterate through each cell in the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // Check for legacy array formula
                    if (cell.IsArrayFormula)
                    {
                        arrayFormulaCount++;
                        continue; // skip double counting if also dynamic (unlikely)
                    }

                    // Check for dynamic array formula
                    if (cell.IsDynamicArrayFormula)
                    {
                        arrayFormulaCount++;
                    }
                }
            }

            // Output the total count
            Console.WriteLine($"Total array formulas in the worksheet: {arrayFormulaCount}");

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("ArrayFormulaCountDemo.xlsx");
        }
    }
}
