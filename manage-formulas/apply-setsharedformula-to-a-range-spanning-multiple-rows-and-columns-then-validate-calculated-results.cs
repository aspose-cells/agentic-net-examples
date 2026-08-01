// Title: C# – Apply SetSharedFormula to a Multi‑Row, Multi‑Column Range and Validate Results with Aspose.Cells
// Description: Creates a workbook, fills A1:A5 with 1‑5, uses SetSharedFormula to assign "=A1*2" to the rectangular range B1:C5, calculates all formulas, prints the computed values, checks that column B equals A × 2 and column C equals B × 2, and saves the file as SharedFormulaResult.xlsx.
// Keywords: Aspose.Cells SetSharedFormula C# | shared formula range | apply shared formula multiple rows | calculate formulas Aspose.Cells | IsSharedFormula property | validate shared formula results | Workbook.CalculateFormula | Excel automation C# | relative reference in shared formula | save workbook Aspose.Cells
// Common Searches: SetSharedFormula for a rectangular range Aspose.Cells | how to validate shared formula values in C# | check IsSharedFormula flag after applying SetSharedFormula | Aspose.Cells calculate formulas after setting shared formula | C# example of shared formula across multiple columns
// Developer Intent: Apply a shared formula to a rectangular block of cells and confirm that the calculated values are correct.
// Use Cases: Populate source data in column A, apply a shared formula to B1:C5, and compute the workbook. | Verify that cells in the shared range report IsSharedFormula = true. | Compare actual cell values with expected multiples to ensure formula accuracy. | Persist the verified workbook by saving it to disk.
// AI Prompts: Generate C# code that uses SetSharedFormula to fill B1:C5 with "=A1*2" and then validates the results against expected values. | Explain how relative references shift when a shared formula is copied to adjacent columns in Aspose.Cells. | Write an MSTest unit test that asserts column B equals column A multiplied by 2 and column C equals column B multiplied by 2 after workbook.CalculateFormula().

using System;
using Aspose.Cells;

namespace AsposeCellsSharedFormulaDemo
{
    // Creates a workbook, fills A1:A5 with 1‑5, uses SetSharedFormula to assign "=A1*2" to the rectangular range B1:C5, calculates all formulas, prints the computed values, checks that column B equals A × 2 and column C equals B × 2, and saves the file as SharedFormulaResult.xlsx.
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // 2. Populate sample data in column A (A1:A5) with values 1 to 5
            for (int row = 0; row < 5; row++)
            {
                cells[row, 0].PutValue(row + 1); // Column index 0 = "A"
            }

            // 3. Set a shared formula starting at cell B1 that will fill a 5x2 range (B1:C5)
            //    The formula "=A1*2" uses a relative reference to the cell on the left.
            //    When the formula is applied to column C, the reference shifts to the left cell (B).
            Cell startCell = cells[0, 1]; // B1
            // Use the overload: SetSharedFormula(string sharedFormula, int rowNumber, int columnNumber)
            startCell.SetSharedFormula("=A1*2", 5, 2);

            // 4. Verify that the cells contain a shared formula
            Console.WriteLine($"B1 IsSharedFormula: {cells[0, 1].IsSharedFormula}");
            Console.WriteLine($"C3 IsSharedFormula: {cells[2, 2].IsSharedFormula}");

            // 5. Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // 6. Output the calculated results for the filled range
            Console.WriteLine("\nCalculated values after applying shared formula (B1:C5):");
            for (int row = 0; row < 5; row++)
            {
                string bValue = cells[row, 1].Value?.ToString() ?? "null"; // Column B
                string cValue = cells[row, 2].Value?.ToString() ?? "null"; // Column C
                Console.WriteLine($"Row {row + 1}: B = {bValue}, C = {cValue}");
            }

            // 7. Simple validation: expected B column = A * 2, C column = B * 2
            bool validationPassed = true;
            for (int row = 0; row < 5; row++)
            {
                double a = Convert.ToDouble(cells[row, 0].Value);
                double expectedB = a * 2;
                double actualB = Convert.ToDouble(cells[row, 1].Value);
                double expectedC = expectedB * 2;
                double actualC = Convert.ToDouble(cells[row, 2].Value);

                if (Math.Abs(expectedB - actualB) > 1e-9 || Math.Abs(expectedC - actualC) > 1e-9)
                {
                    validationPassed = false;
                    Console.WriteLine($"Validation failed at row {row + 1}");
                }
            }
            Console.WriteLine($"\nValidation result: {(validationPassed ? "PASS" : "FAIL")}");

            // 8. Save the workbook (lifecycle rule: use provided save method)
            workbook.Save("SharedFormulaResult.xlsx");
            Console.WriteLine("\nWorkbook saved as 'SharedFormulaResult.xlsx'.");
        }
    }
}
