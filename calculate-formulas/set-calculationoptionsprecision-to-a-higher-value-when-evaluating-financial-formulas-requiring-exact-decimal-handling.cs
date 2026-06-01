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

            // Example financial data
            // A1: Principal
            // A2: Interest Rate
            // A3: Periods
            // A4: Formula = Principal * (1 + Rate) ^ Periods
            cells["A1"].PutValue(1000.00);          // Principal
            cells["A2"].PutValue(0.075);            // Interest Rate (7.5%)
            cells["A3"].PutValue(5);                // Periods (years)

            // Set the compound interest formula
            cells["A4"].Formula = "=A1*(1+A2)^A3";

            // Configure calculation options for higher precision
            CalculationOptions calcOptions = new CalculationOptions
            {
                // Use Decimal strategy to handle calculations with exact decimal precision
                PrecisionStrategy = CalculationPrecisionStrategy.Decimal,

                // Optional: keep other defaults (ignore errors, recursive calculation)
                IgnoreError = true,
                Recursive = true
            };

            // Perform calculation with the specified options
            workbook.CalculateFormula(calcOptions);

            // Retrieve and display the calculated result
            double result = cells["A4"].DoubleValue;
            Console.WriteLine("Compound interest result (high precision): " + result);

            // Save the workbook (you can change the path as needed)
            workbook.Save("FinancialPrecisionDemo.xlsx");
        }
    }
}