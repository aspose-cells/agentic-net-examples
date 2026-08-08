// Title: Implement a Custom Worksheet Function in Aspose.Cells (.NET) Using ICustomFunction and a Custom Engine
// Description: Shows how to create an ICustomFunction implementation, register it with a CustomFunctionDefinition, build a custom calculation engine derived from AbstractCalculationEngine, apply the function in a worksheet formula, run calculation with CalculationOptions, and save the result.
// Keywords: Aspose.Cells custom function .NET | ICustomFunction example | custom calculation engine Aspose | register user‑defined Excel function | array‑mode parameters Aspose.Cells | C# custom worksheet function
// Common Searches: how to add a user defined function in Aspose.Cells | custom calculation engine for Aspose.Cells .NET | define array mode for custom Excel function Aspose | register ICustomFunction implementation | run custom formula with Aspose.Cells
// Developer Intent: Create and integrate a user‑defined worksheet function into Aspose.Cells calculation workflow.
// Use Cases: Add two cell values with a custom MYFUNC and use it in formulas. | Process whole ranges by marking parameters as array‑mode. | Replace the default engine to execute bespoke logic during workbook calculation.
// AI Prompts: Generate C# code for an ICustomFunction that multiplies all numeric arguments and registers it as MYMULT in Aspose.Cells. | Show how to modify MyCustomFunctionDefinition to accept a variable number of arguments and enable array‑mode for the first parameter. | Provide an example that uses the custom engine to evaluate =MYFUNC(A1:A5) where the function sums the range.

using System;
using Aspose.Cells;
using System.Collections.Generic;

// Interface for custom worksheet functions
public interface ICustomFunction
{
    // Evaluates the function with the given arguments
    object Evaluate(object[] args);
}

// Example custom function that sums all numeric arguments
// Shows how to create an ICustomFunction implementation, register it with a CustomFunctionDefinition, build a custom calculation engine derived from AbstractCalculationEngine, apply the function in a worksheet formula, run calculation with CalculationOptions, and save the result.
public class MyCustomFunction : ICustomFunction
{
    public object Evaluate(object[] args)
    {
        double sum = 0;
        foreach (var arg in args)
        {
            if (arg != null && double.TryParse(arg.ToString(), out double d))
                sum += d;
        }
        return sum;
    }
}

// CustomFunctionDefinition to specify which parameters should be calculated in array mode
public class MyCustomFunctionDefinition : CustomFunctionDefinition
{
    public override int[] GetArrayModeParameters(string functionName)
    {
        // For MYFUNC both parameters are processed in array mode
        if (functionName.Equals("MYFUNC", StringComparison.OrdinalIgnoreCase))
            return new int[] { 0, 1 };
        return base.GetArrayModeParameters(functionName);
    }
}

// Custom calculation engine that delegates to registered ICustomFunction implementations
public class MyEngine : AbstractCalculationEngine
{
    private readonly Dictionary<string, ICustomFunction> _functions =
        new Dictionary<string, ICustomFunction>(StringComparer.OrdinalIgnoreCase);

    public MyEngine()
    {
        // Register custom functions here
        _functions["MYFUNC"] = new MyCustomFunction();
    }

    public override void Calculate(CalculationData data)
    {
        if (_functions.TryGetValue(data.FunctionName, out ICustomFunction func))
        {
            // Gather parameter values
            object[] args = new object[data.ParamCount];
            for (int i = 0; i < data.ParamCount; i++)
            {
                args[i] = data.GetParamValue(i);
            }

            // Execute custom function and set the result
            data.CalculatedValue = func.Evaluate(args);
        }
    }
}

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Sample data for the custom function
        ws.Cells["A1"].PutValue(5);
        ws.Cells["A2"].PutValue(7);

        // Update custom function definition (optional, for array‑mode handling)
        wb.UpdateCustomFunctionDefinition(new MyCustomFunctionDefinition());

        // Set a formula that uses the custom function
        ws.Cells["B1"].Formula = "=MYFUNC(A1, A2)";

        // Configure calculation options with the custom engine
        CalculationOptions opts = new CalculationOptions
        {
            CustomEngine = new MyEngine()
        };

        // Perform calculation
        wb.CalculateFormula(opts);

        // Output the result
        Console.WriteLine("Result of MYFUNC: " + ws.Cells["B1"].Value);

        // Save the workbook
        wb.Save("CustomFunctionResult.xlsx");
    }
}
