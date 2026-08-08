// Title: Override TODAY() with a Custom Calculation Engine in Aspose.Cells for .NET
// Description: Demonstrates how to register a custom calculation engine via Workbook.Settings.CustomEngine, intercept the built‑in TODAY() function, return a fixed date, flag its invocation, recalculate the workbook, and verify that all formulas use the overridden implementation.
// Keywords: Aspose.Cells custom engine | override TODAY function | AbstractCalculationEngine C# | Workbook.Settings.CustomEngine | deterministic formula testing | Aspose.Cells .NET example | custom calculation logic | formula interception
// Common Searches: Aspose.Cells replace TODAY() with fixed date | register custom calculation engine Aspose.Cells .NET | how to intercept built‑in functions in Aspose.Cells | verify custom TODAY implementation in workbook | C# Aspose.Cells custom engine tutorial
// Developer Intent: Register a custom calculation engine that overrides TODAY() and confirm that every formula uses this implementation.
// Use Cases: Create repeatable unit tests by fixing TODAY() across all workbook formulas. | Apply organization‑specific date rules by substituting the default TODAY() logic. | Audit or log usage of specific built‑in functions during calculation.
// AI Prompts: Generate C# code that registers a custom calculation engine via Workbook.Settings.CustomEngine to override TODAY() and then recalculates the workbook. | Write a C# unit test using Aspose.Cells that asserts the custom TODAY implementation is called for each formula in a workbook. | Explain how to extend AbstractCalculationEngine to intercept multiple built‑in functions such as TODAY and NOW, and how to expose an Invoked flag for monitoring.

using System;
using Aspose.Cells;

// Demonstrates how to register a custom calculation engine via Workbook.Settings.CustomEngine, intercept the built‑in TODAY() function, return a fixed date, flag its invocation, recalculate the workbook, and verify that all formulas use the overridden implementation.
class CustomTodayEngine : AbstractCalculationEngine
{
    // Indicates whether the custom TODAY implementation was used
    public bool Invoked { get; private set; }

    // Enable processing of built‑in functions so that TODAY can be intercepted
    public override bool ProcessBuiltInFunctions => true;

    public override void Calculate(CalculationData data)
    {
        // Intercept the TODAY function
        if (data.FunctionName.Equals("TODAY", StringComparison.OrdinalIgnoreCase))
        {
            Invoked = true;
            // Return a fixed date for verification purposes
            data.CalculatedValue = new DateTime(2000, 1, 1);
        }
        // For all other functions let the default engine handle them
    }
}

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Place a formula that uses TODAY()
        ws.Cells["A1"].Formula = "=TODAY()";

        // Instantiate the custom calculation engine
        var customEngine = new CustomTodayEngine();

        // Configure calculation options to use the custom engine
        CalculationOptions options = new CalculationOptions
        {
            CustomEngine = customEngine
        };

        // Recalculate all formulas in the workbook using the custom engine
        wb.CalculateFormula(options);

        // Output the calculated value and whether the custom TODAY was invoked
        Console.WriteLine("A1 value (should be 2000‑01‑01): " + ws.Cells["A1"].Value);
        Console.WriteLine("Custom TODAY invoked: " + customEngine.Invoked);

        // Optionally save the workbook to verify the result persists
        wb.Save("CustomTodayResult.xlsx");
    }
}
