using System;
using Aspose.Cells;

namespace AsposeCellsCustomEngineDemo
{
    // Custom engine that overrides TODAY() to return a fixed date (e.g., 2023‑01‑01)
    public class FixedTodayEngine : AbstractCalculationEngine
    {
        // Override Calculate method
        public override void Calculate(CalculationData data)
        {
            // Check if the function being evaluated is TODAY
            if (data.FunctionName != null && data.FunctionName.Equals("TODAY", StringComparison.OrdinalIgnoreCase))
            {
                // Return the fixed date as the calculated value
                data.CalculatedValue = new DateTime(2023, 1, 1);
                // No need to call base.Calculate() because the abstract class has no implementation
            }
            // For all other functions, do nothing and let the default engine handle them
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set a formula that uses TODAY()
            sheet.Cells["A1"].Formula = "=TODAY()";

            // Configure calculation options to use the custom engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new FixedTodayEngine()
            };

            // Calculate all formulas using the custom engine
            workbook.CalculateFormula(options);

            // Output the result of the TODAY() formula (should be the fixed date)
            Console.WriteLine("A1 value (fixed TODAY): " + sheet.Cells["A1"].Value);

            // Save the workbook (demonstrates lifecycle usage)
            workbook.Save("FixedTodayDemo.xlsx");
        }
    }
}