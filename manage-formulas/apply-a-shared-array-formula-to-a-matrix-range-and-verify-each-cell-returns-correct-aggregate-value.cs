// Title: How to apply a shared SUM array formula to a matrix range and validate each cell’s result using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set identical SUM formula for a block of cells | verify that every cell in a range returns the same aggregate sum with Aspose.Cells | apply an array formula to multiple cells in a workbook using Aspose.Cells .NET | calculate and check matrix sum across a range in Aspose.Cells C# example | shared formula across range A5:C7 Aspose.Cells verification
// Tags: shared array formula Aspose.Cells | apply SUM formula to range C# | validate workbook formula results Aspose.Cells | matrix aggregation with Aspose.Cells | formula calculation verification .NET

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example creates a new workbook, fills cells A1:C3 with a 3×3 numeric matrix, computes the expected sum, assigns the formula "=SUM(A1:C3)" to each cell in the target range A5:C7, forces formula calculation, iterates through the range to confirm every cell’s value matches the pre‑computed sum, reports any mismatches, and saves the workbook as SharedArrayFormulaDemo.xlsx.
class SharedArrayFormulaDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Fill a 3x3 matrix (A1:C3) with sample numbers
            int[,] matrix = {
                { 1, 2, 3 },
                { 4, 5, 6 },
                { 7, 8, 9 }
            };

            for (int row = 0; row < matrix.GetLength(0); row++)
            {
                for (int col = 0; col < matrix.GetLength(1); col++)
                {
                    // Cells are zero‑based in Aspose.Cells
                    sheet.Cells[row, col].PutValue(matrix[row, col]);
                }
            }

            // Calculate the expected aggregate value (sum of the matrix)
            double expectedSum = 0;
            foreach (int val in matrix) expectedSum += val;

            // Define the target range where the formula will be applied (A5:C7)
            int startRow = 4; // A5 (zero‑based index)
            int startCol = 0; // column A
            int rows = 3;
            int cols = 3;

            // Apply the same formula to each cell in the target range.
            // Using a regular formula here because the Range class in this version
            // does not expose a FormulaArray property.
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    sheet.Cells[startRow + r, startCol + c].Formula = "=SUM(A1:C3)";
                }
            }

            // Force calculation of all formulas in the workbook
            workbook.CalculateFormula();

            // Verify that each cell in the target range contains the expected sum
            bool allCorrect = true;
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    Cell cell = sheet.Cells[startRow + r, startCol + c];
                    double cellValue = cell.DoubleValue;
                    if (Math.Abs(cellValue - expectedSum) > 1e-9)
                    {
                        allCorrect = false;
                        Console.WriteLine($"Verification failed at {cell.Name}: Expected {expectedSum}, Got {cellValue}");
                    }
                }
            }

            if (allCorrect)
            {
                Console.WriteLine("All cells in the formula range contain the correct aggregate value.");
            }

            // Save the workbook (optional, just to demonstrate lifecycle usage)
            workbook.Save("SharedArrayFormulaDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
