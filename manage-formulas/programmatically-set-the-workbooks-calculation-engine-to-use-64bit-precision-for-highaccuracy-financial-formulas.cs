// Title: Aspose.Cells .NET: Enable 64‑bit Decimal Precision for Accurate Financial Formulas
// Description: Demonstrates how to configure Aspose.Cells calculation engine to use the Decimal precision strategy (64‑bit) in C#. The workbook is created in‑memory, sample values are added, CalculationOptions.PrecisionStrategy is set to Decimal, formulas are recalculated with workbook.CalculateFormula, and the exact result is saved to an Excel file.
// Keywords: Aspose.Cells | C# | .NET | CalculationOptions | PrecisionStrategy | Decimal precision | 64‑bit calculation | high‑accuracy financial formulas | floating point error avoidance | Excel calculation engine
// Common Searches: Aspose.Cells set decimal precision .NET | 64‑bit calculation engine Aspose.Cells | How to avoid floating point errors in Aspose.Cells formulas | CalculationOptions PrecisionStrategy Decimal example | Recalculate workbook with high‑precision financial formulas
// Developer Intent: Configure the workbook to evaluate formulas using 64‑bit decimal arithmetic for maximum numeric accuracy.
// Use Cases: Compute cumulative monetary totals where rounding differences could change the balance. | Validate that the sum of positive and negative amounts returns an exact zero in tax calculations. | Generate interest or amortization schedules that require precise decimal handling across many cells.
// AI Prompts: Show a step‑by‑step C# example that switches Aspose.Cells to Decimal precision and saves the workbook. | Compare the result of a financial sum using the default double precision versus Decimal precision in Aspose.Cells. | Explain how to reset CalculationOptions to the default precision after using the Decimal strategy.

using System;
using Aspose.Cells;

namespace AsposeCellsPrecisionDemo
{
    // Demonstrates how to configure Aspose.Cells calculation engine to use the Decimal precision strategy (64‑bit) in C#. The workbook is created in‑memory, sample values are added, CalculationOptions.PrecisionStrategy is set to Decimal, formulas are recalculated with workbook.CalculateFormula, and the exact result is saved to an Excel file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (in-memory)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Add sample data that may suffer from floating‑point precision issues
            cells["A1"].PutValue(0.45);
            cells["A2"].PutValue(-0.45);
            cells["A3"].PutValue(0.02);
            cells["A4"].Formula = "=A1+A2+A3"; // Expected result: 0

            // Configure calculation options to use 64‑bit decimal precision
            CalculationOptions calcOptions = new CalculationOptions
            {
                // Decimal strategy uses System.Decimal internally for maximum precision
                PrecisionStrategy = CalculationPrecisionStrategy.Decimal
            };

            // Recalculate all formulas with the specified precision strategy
            workbook.CalculateFormula(calcOptions);

            // Output the calculated result to verify high‑precision handling
            Console.WriteLine("Result of A4 with Decimal precision: " + cells["A4"].Value);

            // Save the workbook (demonstrates the required save lifecycle)
            workbook.Save("PrecisionDemo.xlsx");
        }
    }
}
