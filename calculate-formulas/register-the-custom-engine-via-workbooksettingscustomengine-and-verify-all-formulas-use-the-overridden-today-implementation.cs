// Title: Register a custom calculation engine to override the TODAY function in Aspose.Cells for .NET and validate other formulas
// AI Prompts: Implement a subclass of AbstractCalculationEngine that returns a fixed date for the TODAY function, assign it to CalculationOptions.CustomEngine, and recalculate the workbook. | Write verification code that prints the values of cells containing TODAY, SUM, and TODAY+5 to confirm only TODAY is overridden while other functions compute normally.
// Common Searches: asp.net how to replace TODAY function with custom engine in Aspose.Cells | using CalculationOptions.CustomEngine to override specific functions in Aspose.Cells | example of AbstractCalculationEngine for TODAY in Aspose.Cells .NET | validate that custom calculation engine does not affect SUM in Aspose.Cells | register custom engine for workbook formula calculation Aspose.Cells tutorial
// Tags: custom calculation engine Aspose.Cells | override TODAY function Aspose.Cells | CalculationOptions.CustomEngine example | AbstractCalculationEngine TODAY override | formula recalculation with custom engine .NET

using System;
using Aspose.Cells;

// The example creates a workbook, adds formulas using TODAY and SUM, registers a TodayOverrideEngine via CalculationOptions.CustomEngine, recalculates all formulas, and prints the results to show TODAY returns a fixed date (2000‑01‑01) while other functions like SUM work unchanged, then saves the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Add formulas that use TODAY and a regular built‑in function
        ws.Cells["A1"].Formula = "=TODAY()";
        ws.Cells["A2"].Formula = "=SUM(1,2,3)";
        ws.Cells["A3"].Formula = "=TODAY()+5";

        // Register the custom calculation engine that overrides TODAY
        var customEngine = new TodayOverrideEngine();

        // Set calculation options with the custom engine
        CalculationOptions options = new CalculationOptions
        {
            CustomEngine = customEngine
        };

        // Calculate all formulas using the custom engine
        wb.CalculateFormula(options);

        // Verify that TODAY was overridden while other functions work normally
        Console.WriteLine("A1 (TODAY) = " + ws.Cells["A1"].Value); // Expected: 2000-01-01
        Console.WriteLine("A2 (SUM)  = " + ws.Cells["A2"].Value); // Expected: 6
        Console.WriteLine("A3 (TODAY+5) = " + ws.Cells["A3"].Value); // Expected: 2000-01-06

        // Optional: save the workbook to see the results in Excel
        wb.Save("CustomTodayEngine.xlsx");
    }

    // Custom engine that intercepts the TODAY function
    class TodayOverrideEngine : AbstractCalculationEngine
    {
        // Enable processing of built‑in functions so we can handle TODAY
        public override bool ProcessBuiltInFunctions => true;

        public override void Calculate(CalculationData data)
        {
            // Check if the current function is TODAY (case‑insensitive)
            if (string.Equals(data.FunctionName, "TODAY", StringComparison.OrdinalIgnoreCase))
            {
                // Return a fixed date, e.g., 2000‑01‑01
                data.CalculatedValue = new DateTime(2000, 1, 1);
                return;
            }

            // For all other functions let the default engine handle them
            // No action needed here
        }
    }
}
