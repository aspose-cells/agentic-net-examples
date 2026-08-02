// Title: Fallback to Default Calculation Engine in a Custom Aspose.Cells AbstractCalculationEngine (C#)
// Description: Demonstrates a CustomEngine that overrides AbstractCalculationEngine.Calculate to handle a user‑defined function (MYFUNC). For any other function the override leaves CalculatedValue unset, causing Aspose.Cells' built‑in engine to evaluate the formula (e.g., SUM). Includes optional ForceRecalculate logic and a complete runnable example.
// Keywords: Aspose.Cells custom calculation engine | AbstractCalculationEngine Calculate override | fallback to default engine | unsupported function handling | C# custom worksheet function | MYFUNC example | ForceRecalculate Aspose.Cells | Excel formula extension .NET
// Common Searches: Aspose.Cells custom function fallback to built‑in engine | how to handle unknown functions in AbstractCalculationEngine | C# override Calculate for custom Excel functions | Aspose.Cells CalculateFormula with custom engine | leave CalculatedValue unset to trigger default calculation
// Developer Intent: Implement a custom calculation engine that processes specific functions while delegating all other formulas to Aspose.Cells' native engine.
// Use Cases: Add proprietary functions (e.g., MYFUNC) without breaking standard Excel formulas. | Ensure future or misspelled functions are automatically calculated by the default engine. | Force recalculation of custom functions on every workbook calculation.
// AI Prompts: Show code that explicitly calls base.Calculate for unsupported functions in a future version of AbstractCalculationEngine. | Generate a full C# example of a custom AbstractCalculationEngine that supports MYFUNC and falls back to the default engine for all other formulas. | Explain why leaving CalculatedValue unset triggers fallback and how to safely add a base.Calculate call if the abstract class later provides a default implementation.

using System;
using Aspose.Cells;

// Demonstrates a CustomEngine that overrides AbstractCalculationEngine.Calculate to handle a user‑defined function (MYFUNC). For any other function the override leaves CalculatedValue unset, causing Aspose.Cells' built‑in engine to evaluate the formula (e.g., SUM). Includes optional ForceRecalculate logic and a complete runnable example.
public class CustomEngine : AbstractCalculationEngine
{
    // Override Calculate to handle custom functions.
    public override void Calculate(CalculationData data)
    {
        // Example custom function: MYFUNC adds two parameters.
        if (data.FunctionName.Equals("MYFUNC", StringComparison.OrdinalIgnoreCase))
        {
            // Retrieve parameter values (they are already evaluated).
            double p0 = Convert.ToDouble(data.GetParamValue(0));
            double p1 = Convert.ToDouble(data.GetParamValue(1));

            // Set the result for the custom function.
            data.CalculatedValue = p0 + p1;
            return;
        }

        // For any unsupported function, do not set CalculatedValue.
        // Leaving it unset lets the default calculation engine process the function.
        // No explicit call to base.Calculate because the method is abstract.
    }

    // Ensure the custom function is always recalculated (optional).
    public override bool ForceRecalculate(string functionName)
    {
        return functionName.Equals("MYFUNC", StringComparison.OrdinalIgnoreCase);
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new workbook.
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some data.
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["A2"].PutValue(7);

        // Formula using the custom function.
        sheet.Cells["A3"].Formula = "=MYFUNC(A1,A2)";

        // Formula using a built‑in function to demonstrate fallback.
        sheet.Cells["A4"].Formula = "=SUM(A1,A2)";

        // Set calculation options with the custom engine.
        CalculationOptions options = new CalculationOptions
        {
            CustomEngine = new CustomEngine()
        };

        // Perform calculation.
        workbook.CalculateFormula(options);

        // Output results.
        Console.WriteLine("Result of MYFUNC (A3): " + sheet.Cells["A3"].Value);
        Console.WriteLine("Result of SUM (A4): " + sheet.Cells["A4"].Value);

        // Save the workbook (optional).
        workbook.Save("CustomEngineDemo.xlsx");
    }
}
