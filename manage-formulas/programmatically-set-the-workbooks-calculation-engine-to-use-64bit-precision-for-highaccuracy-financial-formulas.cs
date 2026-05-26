using Aspose.Cells;
using System;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Add sample numeric data
        cells["A1"].PutValue(12345.6789);
        cells["A2"].PutValue(9876.54321);

        // Set a formula that will benefit from high‑precision calculation
        cells["A3"].Formula = "=A1+A2";

        // Configure calculation options to use 64‑bit (decimal) precision
        CalculationOptions calcOptions = new CalculationOptions
        {
            // Decimal strategy uses decimal operands for maximum precision
            PrecisionStrategy = CalculationPrecisionStrategy.Decimal
        };

        // Perform calculation with the specified options
        workbook.CalculateFormula(calcOptions);

        // Display the calculated result
        Console.WriteLine("Calculated result with Decimal precision: " + cells["A3"].Value);

        // Save the workbook
        workbook.Save("HighPrecisionWorkbook.xlsx");
    }
}