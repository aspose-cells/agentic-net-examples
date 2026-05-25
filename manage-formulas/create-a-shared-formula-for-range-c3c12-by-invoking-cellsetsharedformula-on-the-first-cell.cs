using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data in column A for rows 3 through 12 (A3:A12)
        for (int row = 2; row < 12; row++) // zero‑based index: row 2 = A3
        {
            cells[row, 0].PutValue(row - 1); // Example values: 2,3,...,11
        }

        // Set a shared formula starting at C3 that will fill C3:C12 (10 rows, 1 column)
        // The formula uses the value from column A of the same row, e.g., =A3*2
        Cell firstCell = cells["C3"];
        firstCell.SetSharedFormula("=A3*2", 10, 1);

        // Recalculate all formulas so that the results are populated
        workbook.CalculateFormula();

        // Output the calculated values for verification
        for (int row = 2; row < 12; row++) // C3:C12
        {
            Console.WriteLine($"C{row + 1} = {cells[row, 2].Value}");
        }

        // Save the workbook
        workbook.Save("SharedFormulaDemo.xlsx");
    }
}