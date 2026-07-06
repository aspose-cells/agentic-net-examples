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