using Aspose.Cells;
using System;

class SharedArrayFormulaDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate a 3x3 matrix (A1:C3) with values 1 through 9
        int val = 1;
        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                cells[row, col].PutValue(val++);
            }
        }

        // Apply a shared formula in column D (D1:D3) that sums each row
        // The formula uses relative references; it will be adjusted for each row automatically
        Cell startCell = cells[0, 3]; // D1
        startCell.SetSharedFormula("=SUM(A1:C1)", 3, 1); // 3 rows, 1 column

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Verify that each cell in D1:D3 contains the correct row sum
        bool allCorrect = true;
        for (int row = 0; row < 3; row++)
        {
            double expectedSum = 0;
            for (int col = 0; col < 3; col++)
            {
                expectedSum += Convert.ToDouble(cells[row, col].Value);
            }

            double actualSum = Convert.ToDouble(cells[row, 3].Value);
            if (Math.Abs(expectedSum - actualSum) > 1e-9)
            {
                allCorrect = false;
            }

            Console.WriteLine($"Row {row + 1} expected sum = {expectedSum}, actual sum = {actualSum}");
        }

        Console.WriteLine("All sums correct: " + allCorrect);

        // Save the workbook (optional)
        workbook.Save("SharedArrayFormulaDemo.xlsx");
    }
}