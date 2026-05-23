using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];

        // Set a formula that uses the TODAY() function
        sheet.Cells["A1"].Formula = "=TODAY()";

        // Create a custom calculation engine that overrides TODAY()
        var customEngine = new TodayOverrideEngine();

        // Register the custom engine via calculation options
        CalculationOptions options = new CalculationOptions
        {
            CustomEngine = customEngine
        };

        // Calculate all formulas using the custom engine
        wb.CalculateFormula(options);

        // Verify that the overridden TODAY value is used
        object result = sheet.Cells["A1"].Value;
        Console.WriteLine("Calculated TODAY value: " + result);

        // Expected custom date (January 1, 2000)
        DateTime expected = new DateTime(2000, 1, 1);
        bool matches = result is DateTime dt && dt.Date == expected.Date;
        Console.WriteLine("Matches expected custom date: " + matches);

        // Save the workbook (optional)
        wb.Save("CustomTodayEngine.xlsx");
    }

    // Custom engine that processes built‑in functions and overrides TODAY()
    class TodayOverrideEngine : AbstractCalculationEngine
    {
        // Indicate that we want to handle built‑in functions ourselves
        public override bool ProcessBuiltInFunctions => true;

        public override void Calculate(CalculationData data)
        {
            // Intercept the TODAY function (case‑insensitive)
            if (string.Equals(data.FunctionName, "TODAY", StringComparison.OrdinalIgnoreCase))
            {
                // Provide a fixed custom date instead of the actual current date
                data.CalculatedValue = new DateTime(2000, 1, 1);
                return;
            }

            // For all other functions, let the default engine handle them
            // (no additional code needed)
        }

        // No special force‑recalculate logic required for this example
        public override bool ForceRecalculate(string functionName) => false;
    }
}