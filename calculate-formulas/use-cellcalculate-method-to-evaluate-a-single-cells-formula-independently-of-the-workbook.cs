// Title: Aspose.Cells C# – Calculate a Single Cell Formula with Cell.Calculate
// Description: Demonstrates how to evaluate only one cell’s formula in an in‑memory workbook using Aspose.Cells for .NET. The example fills A1‑A3, sets a SUM formula in B1, calls Cell.Calculate with CalculationOptions, and returns B1’s result while leaving the rest of the sheet untouched.
// Keywords: Aspose.Cells Cell.Calculate | C# single cell calculation | independent formula evaluation | calculate specific cell .NET | Aspose.Cells example | in‑memory workbook calculation | partial workbook recompute | CalculationOptions usage
// Common Searches: Aspose.Cells calculate only one cell | Cell.Calculate method C# example | evaluate SUM formula without full recalculation | partial calculation Aspose.Cells .NET | how to use CalculationOptions with Cell.Calculate
// Developer Intent: Compute the value of a particular cell’s formula without triggering a full workbook recalculation.
// Use Cases: Retrieve a formula result after updating its referenced cells in a web service. | Run lightweight, on‑the‑fly calculations for user input while keeping the workbook in memory. | Unit‑test a specific formula by calculating only the target cell.
// AI Prompts: Show C# code that uses Cell.Calculate with custom CalculationOptions to evaluate an array formula. | Provide a loop that calculates several independent cells using Cell.Calculate in Aspose.Cells. | Explain how to obtain the calculated value’s data type after calling Cell.Calculate.

using System;
using Aspose.Cells;

namespace AsposeCellsCellCalculateDemo
{
    // Demonstrates how to evaluate only one cell’s formula in an in‑memory workbook using Aspose.Cells for .NET. The example fills A1‑A3, sets a SUM formula in B1, calls Cell.Calculate with CalculationOptions, and returns B1’s result while leaving the rest of the sheet untouched.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (in-memory, no file needed)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Set up some data that the formula will reference
            worksheet.Cells["A1"].PutValue(10);
            worksheet.Cells["A2"].PutValue(20);
            worksheet.Cells["A3"].PutValue(30);

            // Target cell that contains the formula to be evaluated independently
            Cell targetCell = worksheet.Cells["B1"];
            targetCell.Formula = "=SUM(A1:A3)";

            // Calculate only this cell using the Cell.Calculate method
            // Passing a new CalculationOptions instance uses default calculation settings
            targetCell.Calculate(new CalculationOptions());

            // Output the result of the independent calculation
            Console.WriteLine("Result of B1 after independent calculation: " + targetCell.Value);

            // Verify that other cells remain unchanged (they have no formulas)
            Console.WriteLine("A1 value (unchanged): " + worksheet.Cells["A1"].Value);
            Console.WriteLine("A2 value (unchanged): " + worksheet.Cells["A2"].Value);
            Console.WriteLine("A3 value (unchanged): " + worksheet.Cells["A3"].Value);
        }
    }
}
