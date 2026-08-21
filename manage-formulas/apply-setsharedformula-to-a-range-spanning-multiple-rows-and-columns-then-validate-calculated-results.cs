// Title: Apply SetSharedFormula to a Multi‑Row, Multi‑Column Range and Verify Results with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills A1:A5 with 1‑5, sets a shared formula at B1 that multiplies the left cell by 2 across 5 rows and 2 columns, calculates all formulas, prints B1, C1, B5, C5 for verification, and saves the file as SharedFormulaDemo.xlsx.
// Keywords: Aspose.Cells | SetSharedFormula | C# | .NET | shared formula range | calculate formulas | validate formula results | Excel automation | cell formula propagation | workbook calculation
// Common Searches: Aspose.Cells SetSharedFormula example for .NET | How to apply a shared formula to a block of cells in C# | Validate calculated values after using SetSharedFormula | Share one formula across multiple rows and columns in Aspose.Cells | C# code to propagate formulas with SetSharedFormula
// Developer Intent: Apply a single formula to a rectangular cell block and confirm that each cell computes the expected value.
// Use Cases: Generate dependent calculations in adjacent columns after populating a source column. | Reduce file size and improve performance by sharing one formula across large worksheets. | Automated testing of spreadsheet logic by calculating formulas and comparing results to expected values.
// AI Prompts: Provide C# code that uses SetSharedFormula to apply a formula to a 10 × 3 range, then checks every cell against an expected value list. | Show how to log mismatched results when validating shared‑formula output in Aspose.Cells for .NET. | Explain common pitfalls and debugging steps for shared formula propagation errors in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsSharedFormulaDemo
{
    // Creates a workbook, fills A1:A5 with 1‑5, sets a shared formula at B1 that multiplies the left cell by 2 across 5 rows and 2 columns, calculates all formulas, prints B1, C1, B5, C5 for verification, and saves the file as SharedFormulaDemo.xlsx.
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
                cells[row, 0].PutValue(row + 1); // Column 0 = A
            }

            // Set a shared formula starting at B1 that will propagate over 5 rows and 2 columns
            // Formula: each cell will multiply the cell to its left by 2 (relative reference)
            // B1 will have =A1*2, C1 will have =B1*2, B2 will have =A2*2, etc.
            Cell startCell = cells[0, 1]; // B1
            startCell.SetSharedFormula("=A1*2", 5, 2); // rowCount = 5, columnCount = 2

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Validate calculated results
            // Expected values:
            // B1 = 1*2 = 2, C1 = 2*2 = 4
            // B5 = 5*2 = 10, C5 = 10*2 = 20
            Console.WriteLine("Validated Results:");
            Console.WriteLine($"B1 = {cells[0, 1].Value} (expected 2)");
            Console.WriteLine($"C1 = {cells[0, 2].Value} (expected 4)");
            Console.WriteLine($"B5 = {cells[4, 1].Value} (expected 10)");
            Console.WriteLine($"C5 = {cells[4, 2].Value} (expected 20)");

            // Save the workbook to a file
            workbook.Save("SharedFormulaDemo.xlsx");
        }
    }
}
