// Title: Aspose.Cells for .NET – Replace INDIRECT Formulas with Static Values
// Description: Demonstrates how to create a workbook, identify cells that use the volatile INDIRECT function, evaluate each formula, write the result back as a plain value, clear the formula, and save the file using Aspose.Cells in C#.
// Keywords: Aspose.Cells | C# | .NET | replace INDIRECT | volatile formula removal | Excel formula to value | calculate and clear formulas | performance optimization | Excel stability | workbook cleanup
// Common Searches: how to remove INDIRECT formulas with Aspose.Cells | replace volatile Excel functions in C# | convert Excel formulas to values using Aspose.Cells | Aspose.Cells evaluate and clear formulas | C# code to eliminate INDIRECT in workbook
// Developer Intent: Convert every INDIRECT formula in a workbook to its evaluated result, removing volatility while preserving other formulas.
// Use Cases: Prepare a workbook for distribution by turning dynamic references into fixed numbers to improve calculation speed. | Clean up legacy Excel files that rely on INDIRECT, ensuring stable recalculation in automated pipelines. | Export Excel data to CSV or other flat formats after converting all formulas, including INDIRECT, to static values.
// AI Prompts: Write C# code with Aspose.Cells that scans a worksheet, finds cells containing INDIRECT, evaluates them, and replaces the formulas with the computed values. | Show a robust method to replace volatile INDIRECT formulas in large Excel files while leaving non‑volatile formulas untouched. | Suggest an approach to rewrite INDIRECT references as direct cell addresses instead of static values using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsIndirectReplacement
{
    // Demonstrates how to create a workbook, identify cells that use the volatile INDIRECT function, evaluate each formula, write the result back as a plain value, clear the formula, and save the file using Aspose.Cells in C#.
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook (lifecycle rule: create)
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // ------------------------------------------------------------
            // Sample data demonstrating the use of INDIRECT
            // ------------------------------------------------------------
            // A1 holds a reference string "B1"
            cells["A1"].PutValue("B1");
            // B1 holds a numeric value
            cells["B1"].PutValue(123);
            // C1 uses INDIRECT to refer to the cell address stored in A1
            cells["C1"].Formula = "=INDIRECT(A1)";

            // Another example where the reference is a range
            // D1 holds "E1:E3"
            cells["D1"].PutValue("E1:E3");
            cells["E1"].PutValue(10);
            cells["E2"].PutValue(20);
            cells["E3"].PutValue(30);
            // F1 uses INDIRECT to refer to the range in D1 and sums it
            cells["F1"].Formula = "=SUM(INDIRECT(D1))";

            // Calculate all formulas so that dependent values are available
            wb.CalculateFormula();

            // ------------------------------------------------------------
            // Replace all formulas that contain the INDIRECT function
            // ------------------------------------------------------------
            // Iterate through all used cells in the worksheet
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // Process only cells that have a formula containing "INDIRECT"
                    if (cell.IsFormula && cell.Formula.IndexOf("INDIRECT", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // Evaluate the original formula to obtain its current result
                        object evaluatedResult = ws.CalculateFormula(cell.Formula);

                        // Replace the formula with the evaluated result (value only)
                        // This removes the volatile INDIRECT dependency
                        cell.PutValue(evaluatedResult);
                        cell.Formula = string.Empty; // clear the formula text
                    }
                }
            }

            // ------------------------------------------------------------
            // Save the modified workbook (lifecycle rule: save)
            // ------------------------------------------------------------
            wb.Save("Workbook_IndirectReplaced.xlsx");
        }
    }
}
