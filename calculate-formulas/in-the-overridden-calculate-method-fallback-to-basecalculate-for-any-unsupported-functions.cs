// Title: How to create a C# custom Aspose.Cells calculation engine that handles a specific function and falls back to the built‑in engine for all other formulas
// AI Prompts: Write a C# class that inherits from AbstractCalculationEngine, overrides the Calculate method to evaluate a user‑defined function named MYFUNC, and leaves other functions untouched so the default engine processes them. | Show how to attach this custom engine to CalculationOptions, invoke workbook.CalculateFormula, and read the results of both the custom MYFUNC formula and a standard SUM formula.
// Common Searches: C# Aspose.Cells custom calculation engine example with fallback for unsupported functions | How to override Calculate in Aspose.Cells to implement a user‑defined Excel function | Assign a custom AbstractCalculationEngine to CalculationOptions in Aspose.Cells .NET | Use custom function MYFUNC in Aspose.Cells while still using built‑in SUM | Fallback to default calculation engine when custom function not matched Aspose.Cells
// Tags: custom AbstractCalculationEngine implementation C# | fallback to default calculation engine | set custom engine in CalculationOptions | user‑defined Excel function MYFUNC | override Calculate method Aspose.Cells

using System;
using Aspose.Cells;

// The example defines a CustomEngine class that inherits from AbstractCalculationEngine and overrides Calculate to compute a custom function MYFUNC while leaving other functions unchanged, causing Aspose.Cells' built‑in engine to handle them. The engine is assigned to CalculationOptions, the workbook calculates both the custom MYFUNC formula and a regular SUM formula, the results are printed, and the workbook is saved.
class CustomEngine : AbstractCalculationEngine
{
    public override void Calculate(CalculationData data)
    {
        // Handle only the custom function "MYFUNC".
        // For any other function we do nothing, allowing Aspose.Cells
        // to perform its default calculation (fallback behavior).
        if (data.FunctionName != null &&
            data.FunctionName.Equals("MYFUNC", StringComparison.OrdinalIgnoreCase))
        {
            // Example implementation: sum two parameters.
            double param0 = Convert.ToDouble(data.GetParamValue(0));
            double param1 = Convert.ToDouble(data.GetParamValue(1));
            data.CalculatedValue = param0 + param1;
        }
        // No else – leaving CalculatedValue untouched triggers the default engine.
    }

    // No special handling for volatile functions.
    public override bool ForceRecalculate(string functionName) => false;
}

class Program
{
    static void Main()
    {
        // Create a new workbook.
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data.
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["A2"].PutValue(7);

        // Formula using the custom function.
        sheet.Cells["A3"].Formula = "=MYFUNC(A1,A2)";

        // Formula using a built‑in function (should be processed by default engine).
        sheet.Cells["B1"].Formula = "=SUM(A1,A2)";

        // Set calculation options with the custom engine.
        CalculationOptions options = new CalculationOptions
        {
            CustomEngine = new CustomEngine()
        };

        // Calculate all formulas.
        workbook.CalculateFormula(options);

        // Display results.
        Console.WriteLine("Result of MYFUNC (A3): " + sheet.Cells["A3"].Value);
        Console.WriteLine("Result of SUM (B1): " + sheet.Cells["B1"].Value);

        // Save the workbook.
        workbook.Save("CustomEngineDemo.xlsx");
    }
}
