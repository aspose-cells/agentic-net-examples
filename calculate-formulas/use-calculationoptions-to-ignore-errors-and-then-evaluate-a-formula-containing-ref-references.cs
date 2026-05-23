using System;
using Aspose.Cells;

namespace AsposeCellsIgnoreErrorDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some data in column B (which will later be deleted)
            for (int i = 0; i < 5; i++)
            {
                cells[i, 1].PutValue(i + 1); // B1:B5 = 1,2,3,4,5
            }

            // Set a formula that sums the values in column B
            cells["A1"].Formula = "=SUM(B1:B5)";

            // Delete column B – the formula now references a non‑existent range and becomes #REF!
            sheet.Cells.DeleteColumn(1); // Column index 1 corresponds to column B

            // Prepare calculation options to ignore errors during formula evaluation
            CalculationOptions calcOptions = new CalculationOptions
            {
                IgnoreError = true   // Suppress errors such as #REF!
            };

            // Calculate all formulas in the workbook using the specified options
            workbook.CalculateFormula(calcOptions);

            // Output the result of the formula after ignoring the #REF! error
            Console.WriteLine("Result of A1 after ignoring errors: " + cells["A1"].StringValue);

            // Save the workbook (optional, just to demonstrate lifecycle usage)
            workbook.Save("IgnoreErrorResult.xlsx");
        }
    }
}