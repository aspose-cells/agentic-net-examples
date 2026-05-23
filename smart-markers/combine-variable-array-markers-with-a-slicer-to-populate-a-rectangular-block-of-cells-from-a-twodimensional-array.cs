using System;
using Aspose.Cells;

namespace AsposeCellsArraySlicerDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Define a two‑dimensional array of values to populate (3 rows x 3 columns)
            object[][] values = new object[3][];
            values[0] = new object[] { 10, 20, 30 };
            values[1] = new object[] { 40, 50, 60 };
            values[2] = new object[] { 70, 80, 90 };

            // Choose a target cell where the array formula will be placed.
            // The formula will spill into a rectangular block of the same size as the values array.
            Cell targetCell = cells["A1"];

            // Set an array formula with the specified dimensions and pre‑calculated values.
            // The dummy formula "=SEQUENCE(3,3)" generates a 3x3 array; the provided 'values'
            // array overrides the calculated result, effectively populating the block.
            targetCell.SetArrayFormula(
                "=SEQUENCE(3,3)",          // array formula expression
                rowNumber: 3,               // number of rows to populate
                columnNumber: 3,            // number of columns to populate
                options: new FormulaParseOptions(),
                values: values);            // pre‑calculated values for the spill range

            // Calculate the workbook to ensure any dependent formulas are refreshed
            workbook.CalculateFormula();

            // (Optional) Display the populated values in the console for verification
            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    Console.Write(cells[row, col].Value + "\t");
                }
                Console.WriteLine();
            }

            // Save the workbook with the populated rectangular block
            workbook.Save("ArraySlicerResult.xlsx");
        }
    }
}