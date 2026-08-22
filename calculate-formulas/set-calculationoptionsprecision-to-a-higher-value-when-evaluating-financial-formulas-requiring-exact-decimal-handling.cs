// Title: Set CalculationOptions.PrecisionStrategy to Decimal for exact financial formula results in Aspose.Cells (C#)
// AI Prompts: Instantiate a CalculationOptions object, assign PrecisionStrategy = CalculationPrecisionStrategy.Decimal, and recalculate a workbook that contains a compound interest formula. | Illustrate how to enable high‑precision decimal calculations for financial formulas in Aspose.Cells using C# code. | Demonstrate saving a workbook after performing decimal‑precision calculations to retain exact financial values.
// Common Searches: how to use decimal precision with Aspose.Cells CalculationOptions in C# for financial formulas | Aspose.Cells calculation precision strategy Decimal example for compound interest | C# set CalculationOptions.PrecisionStrategy to Decimal to avoid rounding errors in spreadsheets
// Tags: Aspose.Cells calculationoptions decimal precision | C# financial formula exact decimal handling | compound interest spreadsheet calculation Aspose.Cells | set precisionstrategy decimal Aspose.Cells | high precision spreadsheet calculations C#

using System;
using Aspose.Cells;

namespace AsposeCellsPrecisionDemo
{
    // The example creates a workbook, inputs principal, interest rate, and years, defines a compound‑interest formula, configures CalculationOptions to use the Decimal precision strategy, recalculates the formula, outputs the precise result, and saves the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Example financial data
            cells["A1"].PutValue(1234.56);          // Principal
            cells["A2"].PutValue(0.075);            // Interest rate (7.5%)
            cells["A3"].PutValue(5);                // Years

            // Formula that calculates compound interest
            // =A1 * (1 + A2) ^ A3
            cells["B1"].Formula = "=A1*(1+A2)^A3";

            // Configure calculation options to use Decimal precision for exact results
            CalculationOptions calcOptions = new CalculationOptions
            {
                PrecisionStrategy = CalculationPrecisionStrategy.Decimal,
                // Optional: keep other defaults
                IgnoreError = true,
                Recursive = true
            };

            // Perform calculation with the specified options
            workbook.CalculateFormula(calcOptions);

            // Retrieve and display the calculated result
            double result = cells["B1"].DoubleValue;
            Console.WriteLine("Compound interest result (Decimal precision): " + result);

            // Save the workbook (optional)
            workbook.Save("FinancialPrecisionDemo.xlsx");
        }
    }
}
