// Title: Set Aspose.Cells Calculation Engine to Decimal Precision for 10‑12 Decimal Places (C#)
// Description: Demonstrates how to configure Aspose.Cells CalculationOptions.PrecisionStrategy to Decimal, enabling 10‑12 decimal place accuracy for scientific formulas, then calculate, display, and save the workbook.
// Keywords: Aspose.Cells | CalculationOptions | PrecisionStrategy | Decimal precision | high‑precision formulas | scientific calculations .NET | custom calculation precision | 10 decimal places | C#
// Common Searches: Aspose.Cells set decimal precision | How to calculate formulas with 10 decimal places in Aspose.Cells | Increase calculation accuracy Aspose.Cells .NET | CalculationOptions PrecisionStrategy Decimal example | High precision scientific formulas Aspose.Cells
// Developer Intent: Configure Aspose.Cells to use the Decimal precision strategy so formulas are evaluated with 10‑12 decimal places.
// Use Cases: Perform engineering calculations that require 10‑12 decimal places and store results in a workbook. | Generate scientific or financial reports where rounding errors must be minimized. | Batch‑process worksheets with high‑precision formulas before exporting to Excel.
// AI Prompts: Show C# code that sets CalculationOptions.PrecisionStrategy to Decimal and calculates a formula with 10‑12 decimal places using Aspose.Cells. | Explain the impact of the Decimal precision strategy on formula accuracy in Aspose.Cells and how to verify the result. | Provide a step‑by‑step guide to configure high‑precision calculations and save the workbook in Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to configure Aspose.Cells CalculationOptions.PrecisionStrategy to Decimal, enabling 10‑12 decimal place accuracy for scientific formulas, then calculate, display, and save the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Example scientific formula that requires high precision (e.g., 3.1415926535e-11)
        worksheet.Cells["A1"].Formula = "=3.1415926535E-11 * 2";

        // Configure calculation options to use the Decimal precision strategy.
        // This strategy performs calculations using the .NET decimal type,
        // providing up to 28‑29 significant digits, which comfortably covers
        // the requested 10‑12 decimal places for scientific formulas.
        CalculationOptions calcOptions = new CalculationOptions
        {
            PrecisionStrategy = CalculationPrecisionStrategy.Decimal
        };

        // Perform the calculation with the custom precision settings
        workbook.CalculateFormula(calcOptions);

        // Output the calculated result to the console
        Console.WriteLine("Calculated value (A1): " + worksheet.Cells["A1"].Value);

        // Save the workbook (demonstrates usage of the required create/save lifecycle)
        workbook.Save("CustomPrecisionResult.xlsx");
    }
}
