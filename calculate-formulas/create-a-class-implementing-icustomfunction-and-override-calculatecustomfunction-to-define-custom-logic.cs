// Title: Create a custom Excel function in Aspose.Cells (.NET) by implementing ICustomFunction
// Description: Demonstrates how to build a user‑defined Excel function in Aspose.Cells for .NET. A class implements ICustomFunction and overrides CalculateCustomFunction to sum the first two arguments, handling scalar values and range references. The custom engine (derived from AbstractCalculationEngine) registers the function under the name MYCUSTOMFUNC, marks it volatile, and integrates it into workbook calculation.
// Keywords: Aspose.Cells custom function | ICustomFunction .NET | CalculateCustomFunction example | user defined Excel function C# | custom calculation engine Aspose | volatile custom function | ReferredArea handling | MYCUSTOMFUNC
// Common Searches: how to add a user defined function in Aspose.Cells C# | ICustomFunction CalculateCustomFunction tutorial | register custom engine in Aspose.Cells workbook | sum first two parameters custom function Aspose | make custom Excel function volatile Aspose.Cells
// Developer Intent: Implement a class that follows ICustomFunction and provides custom formula logic via CalculateCustomFunction, then register it with a custom calculation engine.
// Use Cases: Calculate a custom sum of the first two arguments in a worksheet formula (e.g., =MYCUSTOMFUNC(A1,B1)). | Extract the first cell value from a range argument (ReferredArea) when used in a custom function. | Ensure the function recalculates on every workbook change by marking it volatile in ForceRecalculate.
// AI Prompts: Write an ICustomFunction that multiplies three parameters and integrates it with a custom engine in Aspose.Cells. | Extend MyCustomEngine to support multiple custom functions identified by distinct names. | Provide a step‑by‑step guide to debug type‑conversion errors inside CalculateCustomFunction.

using System;
using Aspose.Cells;

// Define the interface expected for custom functions
public interface ICustomFunction
{
    // Method that will be called to calculate the custom function
    void CalculateCustomFunction(CalculationData data);
}

// Implementation of a custom function that sums the first two parameters
// Demonstrates how to build a user‑defined Excel function in Aspose.Cells for .NET. A class implements ICustomFunction and overrides CalculateCustomFunction to sum the first two arguments, handling scalar values and range references. The custom engine (derived from AbstractCalculationEngine) registers the function under the name MYCUSTOMFUNC, marks it volatile, and integrates it into workbook calculation.
public class MyCustomFunction : ICustomFunction
{
    public void CalculateCustomFunction(CalculationData data)
    {
        // Ensure we have at least two parameters
        if (data.ParamCount >= 2)
        {
            double sum = 0;

            // Process the first two parameters
            for (int i = 0; i < 2; i++)
            {
                object param = data.GetParamValue(i);

                // If the parameter is a ReferredArea (range), take the first cell value
                if (param is ReferredArea area)
                {
                    sum += Convert.ToDouble(area.GetValue(0, 0));
                }
                else
                {
                    sum += Convert.ToDouble(param);
                }
            }

            // Set the calculated result
            data.CalculatedValue = sum;
        }
        else
        {
            // Not enough parameters – return an error value
            data.CalculatedValue = "#VALUE!";
        }
    }
}

// Custom calculation engine that delegates to ICustomFunction implementations
public class MyCustomEngine : AbstractCalculationEngine
{
    private readonly ICustomFunction _customFunction;

    public MyCustomEngine(ICustomFunction customFunction)
    {
        _customFunction = customFunction;
    }

    public override void Calculate(CalculationData data)
    {
        // Handle only the specific custom function name
        if (string.Equals(data.FunctionName, "MYCUSTOMFUNC", StringComparison.OrdinalIgnoreCase))
        {
            _customFunction.CalculateCustomFunction(data);
        }
        // For all other functions let the default engine handle them
    }

    public override bool ForceRecalculate(string functionName)
    {
        // Ensure the custom function is recalculated for each cell (volatile behavior)
        return string.Equals(functionName, "MYCUSTOMFUNC", StringComparison.OrdinalIgnoreCase);
    }
}

// Demo program
public class Program
{
    public static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["B1"].PutValue(25);

        // Use the custom function in a formula
        sheet.Cells["C1"].Formula = "=MYCUSTOMFUNC(A1,B1)";

        // Instantiate the custom function implementation
        ICustomFunction customFunc = new MyCustomFunction();

        // Set calculation options to use the custom engine
        CalculationOptions options = new CalculationOptions
        {
            CustomEngine = new MyCustomEngine(customFunc)
        };

        // Perform calculation
        workbook.CalculateFormula(options);

        // Output the result
        Console.WriteLine("Result of MYCUSTOMFUNC(A1,B1): " + sheet.Cells["C1"].Value);

        // Save the workbook
        workbook.Save("CustomFunctionResult.xlsx");
    }
}
