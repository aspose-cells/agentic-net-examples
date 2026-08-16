// Title: Enable Decimal Precision in Aspose.Cells C# for Accurate Financial Formula Evaluation
// Description: This example shows how to configure Aspose.Cells CalculationOptions with the Decimal precision strategy, apply it to a workbook containing a compound‑interest formula, recalculate the sheet, and save the result, ensuring exact decimal handling for financial calculations.
// Keywords: Aspose.Cells | C# | CalculationOptions | Decimal precision | PrecisionStrategy | financial formulas | high‑precision calculation | compound interest | tax calculation | Excel formula accuracy
// Common Searches: Aspose.Cells set decimal precision C# | CalculationOptions PrecisionStrategy Decimal example | high precision financial formulas Aspose.Cells | calculate formula with exact decimal handling | compound interest calculation Aspose.Cells
// Developer Intent: Increase formula calculation accuracy by applying the Decimal precision strategy in Aspose.Cells for financial scenarios.
// Use Cases: Compute compound‑interest or loan amortization without rounding errors. | Generate tax worksheets that require exact decimal results. | Recalculate large workbooks with high precision before exporting reports.
// AI Prompts: Write C# code that sets CalculationOptions.PrecisionStrategy to Decimal and recalculates all formulas in an Aspose.Cells workbook. | Explain the impact of Decimal precision on financial calculations in Aspose.Cells and provide a loan schedule example. | Create a step‑by‑step guide to enable high‑precision calculations across multiple worksheets using Aspose.Cells.

using System;
using Aspose.Cells;

// This example shows how to configure Aspose.Cells CalculationOptions with the Decimal precision strategy, apply it to a workbook containing a compound‑interest formula, recalculate the sheet, and save the result, ensuring exact decimal handling for financial calculations.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Sample financial data
        cells["A1"].PutValue(1234.56); // Principal amount
        cells["A2"].PutValue(0.075);   // Annual interest rate (7.5%)
        cells["A3"].PutValue(5);       // Number of years

        // Compound interest formula: =A1*(1+A2)^A3
        cells["B1"].Formula = "=A1*(1+A2)^A3";

        // Configure calculation options to use Decimal precision for exact decimal handling
        CalculationOptions calcOptions = new CalculationOptions
        {
            PrecisionStrategy = CalculationPrecisionStrategy.Decimal
        };

        // Perform calculation with the specified options
        workbook.CalculateFormula(calcOptions);

        // Display the result
        Console.WriteLine("Compound interest result: " + cells["B1"].Value);

        // Save the workbook (optional)
        workbook.Save("FinancialPrecisionDemo.xlsx");
    }
}
