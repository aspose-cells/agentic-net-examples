using System;
using Aspose.Cells;

namespace AsposeCellsCustomEngineDemo
{
    // Custom engine that replaces the TODAY() function with a fixed date.
    public class FixedTodayEngine : AbstractCalculationEngine
    {
        // Override Calculate to handle the TODAY function.
        public override void Calculate(CalculationData data)
        {
            // Check if the function being evaluated is TODAY (case‑insensitive).
            if (string.Equals(data.FunctionName, "TODAY", StringComparison.OrdinalIgnoreCase))
            {
                // Set the calculated value to a fixed date, e.g., 2020‑01‑01.
                data.CalculatedValue = new DateTime(2020, 1, 1);
                // No need to call base.Calculate() because the abstract class has no implementation.
            }
            // For all other functions let the default engine handle them (do nothing here).
        }

        // Ensure TODAY is always recalculated (optional but safe for shared formulas).
        public override bool ForceRecalculate(string functionName)
        {
            return string.Equals(functionName, "TODAY", StringComparison.OrdinalIgnoreCase);
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Place the TODAY() formula in cell A1.
            sheet.Cells["A1"].Formula = "=TODAY()";

            // Set up calculation options with the custom engine.
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new FixedTodayEngine()
            };

            // Calculate all formulas using the custom engine.
            workbook.CalculateFormula(options);

            // Retrieve and display the result of the TODAY() formula.
            object result = sheet.Cells["A1"].Value;
            Console.WriteLine("A1 (TODAY) calculated value: " + result);

            // Save the workbook to demonstrate that the custom calculation persisted.
            workbook.Save("FixedTodayDemo.xlsx");
        }
    }
}