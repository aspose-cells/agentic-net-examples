// Title: C# – Apply a Shared Formula to Sum Rows in a Matrix with Aspose.Cells
// Description: Creates a workbook, fills a 3×2 matrix in A1:B3, assigns a shared formula in D1 that adds the two cells of each row, propagates it to D2‑D3 via SetSharedFormula, calculates the sheet, and programmatically checks that every D‑column value equals the expected row total.
// Keywords: Aspose.Cells | C# | SetSharedFormula | shared formula | row sum | matrix range | calculate formulas | verify results | Excel automation | workbook calculation
// Common Searches: Aspose.Cells SetSharedFormula example | how to sum rows with shared formula in .NET | C# verify Excel formula results programmatically | shared array formula Aspose.Cells | calculate workbook after setting formula
// Developer Intent: Show how to assign a shared formula that computes each row’s total and validate the computed values against expected sums.
// Use Cases: Generate row totals for a data table without writing separate formulas for each cell. | Minimize spreadsheet size by reusing a single formula across many rows. | Automate unit tests that confirm formula accuracy after workbook.CalculateFormula().
// AI Prompts: Write C# code that uses Aspose.Cells SetSharedFormula to sum each row of a given matrix range. | Explain how to validate that every cell containing a shared formula returns the correct sum after workbook.CalculateFormula(). | Provide troubleshooting steps when a shared formula in Aspose.Cells does not produce expected results.

using System;
using Aspose.Cells;

// Creates a workbook, fills a 3×2 matrix in A1:B3, assigns a shared formula in D1 that adds the two cells of each row, propagates it to D2‑D3 via SetSharedFormula, calculates the sheet, and programmatically checks that every D‑column value equals the expected row total.
class SharedArrayFormulaDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate a 3x2 matrix (A1:B3) with sample numeric data
        int[,] matrix = { { 1, 2 }, { 3, 4 }, { 5, 6 } };
        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 2; col++)
            {
                cells[row, col].PutValue(matrix[row, col]); // A1:B3
            }
        }

        // Apply a shared formula that sums each row's A and B values.
        // The formula is entered in D1 and will be propagated to D2 and D3.
        // Parameters: (formula, rowNumber, columnNumber)
        cells["D1"].SetSharedFormula("=SUM(A1:B1)", 3, 1);

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Verify that each cell in D1:D3 contains the correct row sum
        bool allCorrect = true;
        for (int row = 0; row < 3; row++)
        {
            double expected = matrix[row, 0] + matrix[row, 1];
            double actual = cells[row, 3].DoubleValue; // Column D has index 3
            Console.WriteLine($"Row {row + 1}: Expected = {expected}, Actual = {actual}");
            if (Math.Abs(expected - actual) > 1e-9)
                allCorrect = false;
        }

        Console.WriteLine("All values correct: " + allCorrect);
    }
}
