using System;
using Aspose.Cells;

namespace AsposeCellsIgnoreErrorDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Set a formula that will cause division by zero
            cells["A1"].Formula = "=1/0";

            // Set another normal formula for comparison
            cells["A2"].Formula = "=10/2";

            // Configure calculation options to ignore errors (e.g., division by zero)
            CalculationOptions calcOptions = new CalculationOptions
            {
                IgnoreError = true   // Suppress errors during calculation
            };

            // Perform calculation with the specified options
            workbook.CalculateFormula(calcOptions);

            // Output the results
            Console.WriteLine("A1 (division by zero, error ignored): " + cells["A1"].StringValue);
            Console.WriteLine("A2 (normal calculation): " + cells["A2"].StringValue);
        }
    }
}