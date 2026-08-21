// Title: Shared Array Formula (SetArrayFormula) – Sum Matrix Verification in Aspose.Cells for .NET
// Description: Creates a 3×3 numeric matrix (A1:C3), defines =SUM($A$1:$C$3) as a shared array, applies it to B2:D4 with SetArrayFormula, calculates all formulas, retrieves the expected total via a regular SUM, and confirms each spilled cell matches the aggregate before saving the workbook.
// Keywords: Aspose.Cells | C# | SetArrayFormula | shared array formula | matrix sum | verify array result | calculate formulas | Excel automation .NET | spilled array range | Workbook.Save
// Common Searches: Aspose.Cells SetArrayFormula example C# | how to apply shared array formula to a range | verify spilled array values in Aspose.Cells | calculate sum of matrix with Aspose.Cells | C# code for shared array formula verification
// Developer Intent: Apply a shared array formula to a matrix range and ensure every cell returns the correct aggregate sum.
// Use Cases: Display the total of a data block in every cell of a target area using a single shared formula. | Programmatically validate that a spilled array produces identical results across all cells. | Generate an Excel file where the summed value is visible in a multi‑cell block for reporting or further processing.
// AI Prompts: Modify the example to use AVERAGE instead of SUM while keeping the verification steps. | Show how to apply a shared array formula to a non‑contiguous range with Aspose.Cells for .NET. | Explain error handling when the source matrix contains text or empty cells during SetArrayFormula execution.

using System;
using Aspose.Cells;

namespace AsposeCellsSharedArrayFormulaDemo
{
    // Creates a 3×3 numeric matrix (A1:C3), defines =SUM($A$1:$C$3) as a shared array, applies it to B2:D4 with SetArrayFormula, calculates all formulas, retrieves the expected total via a regular SUM, and confirms each spilled cell matches the aggregate before saving the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate a 3x3 matrix in A1:C3
            int[,] data = {
                { 1, 2, 3 },
                { 4, 5, 6 },
                { 7, 8, 9 }
            };

            for (int r = 0; r < 3; r++)
                for (int c = 0; c < 3; c++)
                    cells[r, c].PutValue(data[r, c]);

            // Define the array formula that calculates the total sum of the matrix
            string arrayFormula = "=SUM($A$1:$C$3)";

            // Apply the array formula as a shared array to a 3x3 range starting at B2 (cells B2:D4)
            // This will cause the same aggregate value to appear in every cell of the range
            Cell startCell = cells["B2"];
            startCell.SetArrayFormula(arrayFormula, 3, 3, new FormulaParseOptions());

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Get the expected aggregate value using a normal formula calculation
            object expectedObj = sheet.CalculateFormula("=SUM(A1:C3)");
            double expected = Convert.ToDouble(expectedObj);

            // Verify each cell in the spilled range contains the expected aggregate value
            Console.WriteLine("Verifying shared array formula results (B2:D4):");
            for (int r = 0; r < 3; r++)
            {
                for (int c = 0; c < 3; c++)
                {
                    Cell cur = cells[1 + r, 1 + c]; // B2 is (1,1)
                    double actual = Convert.ToDouble(cur.Value);
                    bool match = Math.Abs(actual - expected) < 1e-9;
                    Console.WriteLine($"{cur.Name}: {actual} {(match ? "OK" : "FAIL")}");
                }
            }

            // Save the workbook (optional, just to visualize the result if needed)
            workbook.Save("SharedArrayFormulaDemo.xlsx");
        }
    }
}
