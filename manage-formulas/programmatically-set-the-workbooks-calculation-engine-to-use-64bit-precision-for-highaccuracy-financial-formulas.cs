// Title: Aspose.Cells C# – Enable 64‑bit Decimal Precision for Financial Calculations
// Description: Shows how to configure an Aspose.Cells workbook to use 64‑bit decimal arithmetic via CalculationOptions.PrecisionStrategy, run the calculation engine, and save the file, eliminating floating‑point drift in monetary formulas.
// Keywords: Aspose.Cells | C# | 64-bit precision | Decimal precision | CalculationOptions | PrecisionStrategy | financial formulas | high‑accuracy calculations | Workbook.CalculateFormula | floating point errors
// Common Searches: Aspose.Cells set decimal precision C# | 64-bit calculation engine Aspose.Cells | prevent rounding errors in Aspose.Cells formulas | CalculationOptions PrecisionStrategy example | financial spreadsheet precision Aspose.Cells
// Developer Intent: Set the workbook’s calculation engine to 64‑bit decimal mode so that monetary formulas return exact values.
// Use Cases: Compute interest accruals where tiny rounding differences can change the final balance. | Validate reconciliation reports that require a true zero sum after multiple additions and subtractions. | Generate tax depreciation schedules that depend on precise decimal arithmetic across many periods.
// AI Prompts: Provide a step‑by‑step guide to switch Aspose.Cells to 128‑bit precision for ultra‑high‑accuracy scenarios. | Show an example of applying CalculationOptions with Decimal precision to a workbook containing dozens of complex financial sheets. | Explain the performance trade‑offs of using Decimal precision in Aspose.Cells and how to enable or disable it at runtime.

using Aspose.Cells;
using System;

// Shows how to configure an Aspose.Cells workbook to use 64‑bit decimal arithmetic via CalculationOptions.PrecisionStrategy, run the calculation engine, and save the file, eliminating floating‑point drift in monetary formulas.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Sample data that can expose precision issues
        cells["A1"].PutValue(0.45);
        cells["A2"].PutValue(-0.45);
        // Formula that should result in zero
        cells["A3"].Formula = "=A1+A2";

        // Configure calculation options to use 64‑bit (decimal) precision
        CalculationOptions calcOptions = new CalculationOptions
        {
            PrecisionStrategy = CalculationPrecisionStrategy.Decimal
        };

        // Perform calculation with the specified precision strategy
        workbook.CalculateFormula(calcOptions);

        // Display the calculated result
        Console.WriteLine("A3 result with Decimal precision: " + cells["A3"].Value);

        // Save the workbook
        workbook.Save("HighPrecision.xlsx");
    }
}
