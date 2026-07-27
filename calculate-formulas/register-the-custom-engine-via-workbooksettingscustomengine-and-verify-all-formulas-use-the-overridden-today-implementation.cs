using System;
using Aspose.Cells;

namespace CustomTodayEngineDemo
{
    // Custom calculation engine that overrides the built‑in TODAY function
    public class TodayOverrideEngine : AbstractCalculationEngine
    {
        // Indicate that we want to process built‑in functions ourselves
        public override bool ProcessBuiltInFunctions => true;

        // Called for each function during calculation
        public override void Calculate(CalculationData data)
        {
            // Check for the TODAY function (case‑insensitive)
            if (data.FunctionName.Equals("TODAY", StringComparison.OrdinalIgnoreCase))
            {
                // Return a fixed date instead of the actual current date
                data.CalculatedValue = new DateTime(2020, 1, 1);
                return;
            }

            // For all other functions let the default engine handle them
            // (no action needed because ProcessBuiltInFunctions is false for them)
        }

        // Ensure TODAY is always recalculated (optional but safe for volatile functions)
        public override bool ForceRecalculate(string functionName)
        {
            return functionName.Equals("TODAY", StringComparison.OrdinalIgnoreCase);
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];

            // Add formulas that use TODAY
            sheet.Cells["A1"].Formula = "=TODAY()";
            sheet.Cells["A2"].Formula = "=TODAY() + 5"; // TODAY as a serial number plus 5 days

            // Set up calculation options with the custom engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new TodayOverrideEngine()
            };

            // Calculate all formulas using the custom engine
            wb.CalculateFormula(options);

            // Verify that the overridden TODAY value is used
            Console.WriteLine("A1 (overridden TODAY): " + sheet.Cells["A1"].Value); // Expected 2020‑01‑01
            Console.WriteLine("A2 (TODAY + 5 days): " + sheet.Cells["A2"].Value); // Expected 2020‑01‑06

            // Save the workbook (optional)
            wb.Save("CustomTodayEngineDemo.xlsx");
        }
    }
}