using System;
using Aspose.Cells;

namespace AsposeCellsPrecisionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate cells with financial values (example)
            cells["A1"].PutValue(1234.56);   // Principal
            cells["A2"].PutValue(0.075);     // Interest rate (7.5%)
            cells["A3"].PutValue(5);         // Years

            // Formula that calculates compound interest
            // =A1 * (1 + A2) ^ A3
            cells["B1"].Formula = "=A1*(1+A2)^A3";

            // Configure calculation options to use Decimal precision for exact results
            CalculationOptions calcOptions = new CalculationOptions
            {
                PrecisionStrategy = CalculationPrecisionStrategy.Decimal,
                // Optional: keep other defaults (IgnoreError, Recursive, etc.)
                IgnoreError = true,
                Recursive = true
            };

            // Perform calculation with the specified options
            workbook.CalculateFormula(calcOptions);

            // Retrieve and display the calculated result
            double result = cells["B1"].DoubleValue;
            Console.WriteLine("Compound interest result (Decimal precision): " + result);

            // Save the workbook to verify the result in the file
            workbook.Save("FinancialPrecisionDemo.xlsx");
        }
    }
}