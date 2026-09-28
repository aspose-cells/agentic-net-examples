// Title: Create a custom calculation engine in Aspose.Cells for .NET to replace undefined Excel functions with user‑defined SUM and AVERAGE logic
// AI Prompts: Implement an AbstractCalculationEngine subclass that maps unknown function names (e.g., FOO, BAR) to custom handlers and assign it through CalculationOptions.CustomEngine to recalculate a workbook. | Add a new custom function called MAXX to the SubstituteEngine that returns the maximum of its arguments, then use =MAXX(A1,B2) in a worksheet and display the result. | Demonstrate how to extract cell values from a ReferredArea inside a custom calculation engine to support range arguments for SUM‑like and AVERAGE‑like functions. | Show how to handle unrecognized functions by returning the #NAME? error from the custom engine.
// Common Searches: aspnet how to handle unknown Excel functions with a custom calculation engine in Aspose.Cells | replace custom function FOO with SUM logic using Aspose.Cells C# | example of AbstractCalculationEngine overriding formula evaluation in Aspose.Cells | use CalculationOptions.CustomEngine to map undefined functions to built‑in behavior | retrieve range values inside custom Aspose.Cells calculation engine
// Tags: custom calculation engine Aspose.Cells | user‑defined function substitution .NET | AbstractCalculationEngine handler mapping | override Excel formula evaluation with C# | replace undefined functions with SUM logic | average function implementation in custom engine

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The sample creates a workbook, writes sample data, and assigns formulas that reference undefined functions (FOO, BAR). A SubstituteEngine derived from AbstractCalculationEngine maps these names to SUM‑like and AVERAGE‑like handlers. The engine is set via CalculationOptions.CustomEngine, the workbook is recalculated, results are printed, and the file is saved.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Populate some sample data
        ws.Cells["A1"].PutValue(10);
        ws.Cells["A2"].PutValue(20);
        ws.Cells["A3"].PutValue(30);
        ws.Cells["B1"].PutValue(5);
        ws.Cells["B2"].PutValue(15);

        // Formulas that use undefined functions (FOO, BAR)
        // These will be substituted by the custom engine
        ws.Cells["C1"].Formula = "=FOO(A1,A2)"; // should behave like SUM
        ws.Cells["C2"].Formula = "=BAR(B1,B2)"; // should behave like AVERAGE

        // Set calculation options to use the custom engine
        CalculationOptions opts = new CalculationOptions
        {
            CustomEngine = new SubstituteEngine()
        };

        // Calculate all formulas using the custom engine
        wb.CalculateFormula(opts);

        // Display the results
        Console.WriteLine("C1 (FOO) = " + ws.Cells["C1"].Value);
        Console.WriteLine("C2 (BAR) = " + ws.Cells["C2"].Value);

        // Save the workbook
        wb.Save("SubstituteEngineDemo.xlsx");
    }
}

// Custom calculation engine that substitutes missing functions with user‑defined equivalents
class SubstituteEngine : AbstractCalculationEngine
{
    // Mapping from unknown function name to a handler that performs the calculation
    private readonly Dictionary<string, Func<CalculationData, object>> _substitutes;

    public SubstituteEngine()
    {
        _substitutes = new Dictionary<string, Func<CalculationData, object>>(StringComparer.OrdinalIgnoreCase)
        {
            { "FOO", SumFunction },      // FOO will act like SUM
            { "BAR", AverageFunction }   // BAR will act like AVERAGE
        };
    }

    // Core method called by Aspose.Cells for each function occurrence
    public override void Calculate(CalculationData data)
    {
        if (_substitutes.TryGetValue(data.FunctionName, out var handler))
        {
            // Use the mapped handler to compute the result
            data.CalculatedValue = handler(data);
        }
        else
        {
            // Function not recognized – return #NAME? so Excel shows an error
            data.CalculatedValue = "#NAME?";
        }
    }

    // Implementation of a SUM‑like function
    private object SumFunction(CalculationData data)
    {
        double sum = 0;
        for (int i = 0; i < data.ParamCount; i++)
        {
            object param = data.GetParamValue(i);
            if (param is ReferredArea area)
            {
                // Iterate over the range
                for (int r = area.StartRow; r <= area.EndRow; r++)
                {
                    for (int c = area.StartColumn; c <= area.EndColumn; c++)
                    {
                        object cellVal = area.GetValue(r - area.StartRow, c - area.StartColumn);
                        sum += Convert.ToDouble(cellVal);
                    }
                }
            }
            else
            {
                sum += Convert.ToDouble(param);
            }
        }
        return sum;
    }

    // Implementation of an AVERAGE‑like function
    private object AverageFunction(CalculationData data)
    {
        double sum = 0;
        int count = 0;
        for (int i = 0; i < data.ParamCount; i++)
        {
            object param = data.GetParamValue(i);
            if (param is ReferredArea area)
            {
                for (int r = area.StartRow; r <= area.EndRow; r++)
                {
                    for (int c = area.StartColumn; c <= area.EndColumn; c++)
                    {
                        object cellVal = area.GetValue(r - area.StartRow, c - area.StartColumn);
                        sum += Convert.ToDouble(cellVal);
                        count++;
                    }
                }
            }
            else
            {
                sum += Convert.ToDouble(param);
                count++;
            }
        }
        return count == 0 ? 0 : sum / count;
    }

    // No need to force recalculation for these functions
    public override bool ForceRecalculate(string functionName) => false;
}
