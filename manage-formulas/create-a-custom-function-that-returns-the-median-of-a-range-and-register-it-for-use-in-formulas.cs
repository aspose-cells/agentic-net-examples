// Title: Create & Register a Custom MEDIAN Function in Aspose.Cells for .NET
// Description: Demonstrates how to build a custom calculation engine that computes the median of numeric values, define the MEDIAN function with array‑mode parameters, register it in a workbook, use the formula "=MEDIAN(A1:A5)", and save the result with Aspose.Cells for C#.
// Keywords: Aspose.Cells custom function | C# median function | .NET custom calculation engine | array mode parameters | register custom function definition | worksheet formula extension | calculate median range | Aspose.Cells example | custom MEDIAN implementation
// Common Searches: Aspose.Cells custom median function example | how to register a custom function in Aspose.Cells .NET | array mode parameter in Aspose.Cells custom function | calculate median of a range with Aspose.Cells | custom calculation engine Aspose.Cells C#
// Developer Intent: Add a user‑defined MEDIAN function to an Aspose.Cells workbook and make it callable from standard formulas.
// Use Cases: Compute the median of a column or row of numbers directly in a worksheet formula. | Support single‑cell or multi‑cell arguments while ignoring non‑numeric entries. | Extend Aspose.Cells with additional statistical functions without modifying the core library.
// AI Prompts: Generate C# code that creates a custom Aspose.Cells MEDIAN function, registers it, and uses it in a worksheet formula. | Explain why array‑mode parameters are required for range‑based custom functions in Aspose.Cells. | Write a unit test in C# that validates the custom MEDIAN function returns correct results for mixed numeric and empty cells.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Custom calculation engine that computes the median of a range
// Demonstrates how to build a custom calculation engine that computes the median of numeric values, define the MEDIAN function with array‑mode parameters, register it in a workbook, use the formula "=MEDIAN(A1:A5)", and save the result with Aspose.Cells for C#.
class MyMedianEngine : AbstractCalculationEngine
{
    public override void Calculate(CalculationData data)
    {
        // Handle only the custom MEDIAN function
        if (data.FunctionName.Equals("MEDIAN", StringComparison.OrdinalIgnoreCase))
        {
            // Retrieve the first parameter (expected to be in array mode)
            object param = data.GetParamValue(0);
            double[] values;

            // If the parameter is an array, extract numeric values
            if (param is object[,] arr)
            {
                List<double> list = new List<double>();
                foreach (var v in arr)
                {
                    if (v != null && double.TryParse(v.ToString(), out double d))
                        list.Add(d);
                }
                values = list.ToArray();
            }
            else
            {
                // Single value case
                if (double.TryParse(param?.ToString() ?? "0", out double d))
                    values = new double[] { d };
                else
                    values = new double[0];
            }

            // If no numeric values, return 0
            if (values.Length == 0)
            {
                data.CalculatedValue = 0;
                return;
            }

            // Sort and compute median
            Array.Sort(values);
            int n = values.Length;
            double median = (n % 2 == 1)
                ? values[n / 2]
                : (values[n / 2 - 1] + values[n / 2]) / 2.0;

            data.CalculatedValue = median;
        }
    }
}

// Custom function definition that marks the first parameter of MEDIAN as array‑mode
class MyCustomFunctionDefinition : CustomFunctionDefinition
{
    public override int[] GetArrayModeParameters(string functionName)
    {
        if (functionName.Equals("MEDIAN", StringComparison.OrdinalIgnoreCase))
            return new int[] { 0 }; // first parameter needs array mode
        return base.GetArrayModeParameters(functionName);
    }
}

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Populate sample data in A1:A5
        ws.Cells["A1"].PutValue(5);
        ws.Cells["A2"].PutValue(2);
        ws.Cells["A3"].PutValue(9);
        ws.Cells["A4"].PutValue(4);
        ws.Cells["A5"].PutValue(7);

        // Register the custom function definition (so MEDIAN's parameter is array mode)
        wb.UpdateCustomFunctionDefinition(new MyCustomFunctionDefinition());

        // Set a formula that uses the custom MEDIAN function
        ws.Cells["B1"].Formula = "=MEDIAN(A1:A5)";

        // Calculate formulas using the custom engine
        CalculationOptions calcOpts = new CalculationOptions
        {
            CustomEngine = new MyMedianEngine()
        };
        wb.CalculateFormula(calcOpts);

        // Output the result
        Console.WriteLine("Median of A1:A5 = " + ws.Cells["B1"].Value);

        // Save the workbook (lifecycle rule)
        wb.Save("MedianCustomFunction.xlsx");
    }
}
