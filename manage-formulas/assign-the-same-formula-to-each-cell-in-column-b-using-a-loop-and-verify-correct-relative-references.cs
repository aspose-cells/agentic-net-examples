using System;
using Aspose.Cells;

namespace AsposeCellsFormulaDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data in column A (A1:A5) with values 1 to 5
            for (int row = 0; row < 5; row++)
            {
                cells[row, 0].PutValue(row + 1); // Column index 0 = A
            }

            // Assign the same relative formula to each cell in column B (B1:B5)
            // Formula references the cell in column A of the same row, e.g., B1 = A1*2
            for (int row = 0; row < 5; row++)
            {
                // Using the Formula property (A1‑style) to set the expression
                cells[row, 1].Formula = $"=A{row + 1}*2"; // Column index 1 = B
            }

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Verify that the formulas produced the expected results
            for (int row = 0; row < 5; row++)
            {
                double expected = (row + 1) * 2;
                Console.WriteLine($"B{row + 1} value: {cells[row, 1].Value} (expected {expected})");
            }

            // Save the workbook to a file
            workbook.Save("ColumnBFormulaDemo.xlsx");
        }
    }
}