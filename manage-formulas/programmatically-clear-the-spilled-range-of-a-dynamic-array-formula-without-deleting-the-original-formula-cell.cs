using System;
using Aspose.Cells;

namespace AsposeCellsDynamicArrayClearSpill
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Set a dynamic array formula in cell A1 (will spill to a 3x2 range)
            Cell formulaCell = cells["A1"];
            formulaCell.SetDynamicArrayFormula("=SEQUENCE(3,2)", new FormulaParseOptions(), true);

            // Calculate formulas so the spill range is populated
            workbook.CalculateFormula();

            // Retrieve the full spill range of the dynamic array formula
            CellArea spillRange = formulaCell.GetArrayRange();

            // If the spill range consists of more than the formula cell, clear the spilled cells
            if (spillRange.StartRow != spillRange.EndRow || spillRange.StartColumn != spillRange.EndColumn)
            {
                // Clear cells to the right of the formula cell in the first row (if any)
                if (spillRange.StartColumn < spillRange.EndColumn)
                {
                    cells.ClearContents(
                        spillRange.StartRow,                     // start row (first row)
                        spillRange.StartColumn + 1,              // start column (one column right of formula)
                        spillRange.StartRow,                     // end row (still first row)
                        spillRange.EndColumn);                   // end column (last column of spill)
                }

                // Clear cells below the first row (entire width) (if any)
                if (spillRange.StartRow < spillRange.EndRow)
                {
                    cells.ClearContents(
                        spillRange.StartRow + 1,                 // start row (one row below formula)
                        spillRange.StartColumn,                  // start column (first column of spill)
                        spillRange.EndRow,                       // end row (last row of spill)
                        spillRange.EndColumn);                   // end column (last column of spill)
                }
            }

            // At this point, the original formula cell (A1) still contains the formula,
            // while all other cells of the spilled range have been cleared.

            // Save the workbook to verify the result (optional)
            workbook.Save("DynamicArraySpillCleared.xlsx");
        }
    }
}