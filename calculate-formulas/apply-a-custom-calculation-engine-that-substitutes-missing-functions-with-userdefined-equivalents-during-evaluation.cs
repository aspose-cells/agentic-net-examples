// Title: How to create and register a custom worksheet function with a custom calculation engine in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# custom function MYFUNC sum range example | register custom calculation engine in Aspose.Cells .NET to handle unknown Excel functions | use CustomFunctionDefinition and AbstractCalculationEngine to evaluate user‑defined formulas in Aspose.Cells | calculate workbook formulas with a custom engine while ignoring errors in Aspose.Cells C# | set array‑mode parameter for a custom function in Aspose.Cells calculation options
// Tags: custom function definition Aspose.Cells | custom calculation engine Aspose.Cells | array mode parameter Aspose.Cells | evaluate missing Excel functions C# | workbook.CalculateFormula custom engine

using System;
using Aspose.Cells;

namespace CustomCalculationEngineDemo
{
    // The example demonstrates how to define a user‑defined function (MYFUNC) that sums all numbers in a range, register its definition to treat the first argument in array mode, implement a custom calculation engine by extending AbstractCalculationEngine, plug the engine into CalculationOptions, execute workbook.CalculateFormula, and output the computed result (100) using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data that the custom function will use
            // A1:B2 matrix:
            // 10  20
            // 30  40
            cells["A1"].PutValue(10);
            cells["B1"].PutValue(20);
            cells["A2"].PutValue(30);
            cells["B2"].PutValue(40);

            // Set a formula that uses a missing (custom) function MYFUNC
            // The function is expected to sum all numbers in the supplied range.
            cells["C1"].Formula = "=MYFUNC(A1:B2)";

            // Create and register a custom function definition.
            // This tells the calculation engine which parameters of MYFUNC should be
            // evaluated in array mode (the first parameter in this case).
            workbook.UpdateCustomFunctionDefinition(new MyCustomFunctionDefinition());

            // Create a custom calculation engine that knows how to evaluate MYFUNC.
            AbstractCalculationEngine customEngine = new MyCustomCalculationEngine();

            // Prepare calculation options and plug in the custom engine.
            CalculationOptions calcOptions = new CalculationOptions
            {
                CustomEngine = customEngine,
                // Optional: ignore errors for any other unsupported functions
                IgnoreError = true
            };

            // Perform the calculation using the custom engine.
            workbook.CalculateFormula(calcOptions);

            // Output the result of the custom function.
            Console.WriteLine("Result of MYFUNC(A1:B2): " + cells["C1"].Value);
            // Expected result: 10+20+30+40 = 100
        }
    }

    // CustomFunctionDefinition allows us to specify which parameters of a custom
    // function should be processed in array mode. Here we indicate that the first
    // (index 0) parameter of MYFUNC is an array.
    public class MyCustomFunctionDefinition : CustomFunctionDefinition
    {
        public override int[] GetArrayModeParameters(string functionName)
        {
            // If the function name matches our custom function, return the index
            // of parameters that need array‑mode evaluation.
            if (functionName.Equals("MYFUNC", StringComparison.OrdinalIgnoreCase))
            {
                // Parameter at index 0 (the first parameter) should be treated as an array.
                return new int[] { 0 };
            }

            // For any other function, no array‑mode parameters.
            return Array.Empty<int>();
        }
    }

    // The custom calculation engine that actually computes the value of MYFUNC.
    // It derives from AbstractCalculationEngine and overrides the Calculate method.
    public class MyCustomCalculationEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // Ensure we are handling the correct custom function.
            if (!data.FunctionName.Equals("MYFUNC", StringComparison.OrdinalIgnoreCase))
                return; // Let the default engine handle other functions.

            // MYFUNC is defined to sum all numeric values in the supplied range.
            // Because we marked the first parameter as array‑mode, GetParamValue(0)
            // will return a two‑dimensional object array representing the range.
            object param = data.GetParamValue(0);

            double sum = 0;

            if (param is object[,] array)
            {
                int rows = array.GetLength(0);
                int cols = array.GetLength(1);
                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        if (array[r, c] is double d)
                        {
                            sum += d;
                        }
                        else if (array[r, c] is int i)
                        {
                            sum += i;
                        }
                        else if (double.TryParse(Convert.ToString(array[r, c]), out double parsed))
                        {
                            sum += parsed;
                        }
                    }
                }
            }
            else if (param is double d)
            {
                sum = d;
            }
            else if (param is int i)
            {
                sum = i;
            }
            else
            {
                // If the parameter is not a range or numeric, treat result as 0.
                sum = 0;
            }

            // Assign the calculated result back to the engine.
            data.CalculatedValue = sum;
        }
    }
}
