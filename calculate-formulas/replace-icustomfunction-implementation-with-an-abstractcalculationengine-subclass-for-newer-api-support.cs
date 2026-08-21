// Title: Implement a custom calculation engine with AbstractCalculationEngine for a user‑defined MYFUNC in Aspose.Cells for .NET
// AI Prompts: Create a C# class that inherits from AbstractCalculationEngine and overrides Calculate to sum two ReferredArea parameters for the function MYFUNC. | Override ForceRecalculate in the custom engine so that only the MYFUNC function forces a recalculation. | Configure CalculationOptions.CustomEngine with the new engine, assign a formula using MYFUNC, and call Workbook.CalculateFormula to obtain the result.
// Common Searches: how to replace ICustomFunction with AbstractCalculationEngine in Aspose.Cells | example of custom function MYFUNC using AbstractCalculationEngine C# | using ReferredArea objects in Aspose.Cells custom calculation engine | configure CalculationOptions to use a custom AbstractCalculationEngine in .NET | force recalculation for specific custom function in Aspose.Cells
// Tags: abstractcalculationengine subclass for user-defined functions | aspocells custom function using ReferredArea | calculationoptions set custom engine | force-recalculate specific function Aspose.Cells | sum two cell ranges in custom Aspose.Cells engine

using System;
using Aspose.Cells;

namespace AsposeCellsCustomEngineDemo
{
    // Custom calculation engine that replaces the old ICustomFunction implementation
    // Demonstrates how to replace the legacy ICustomFunction with an AbstractCalculationEngine subclass, override ForceRecalculate and Calculate to handle a user‑defined function MYFUNC that sums two cell ranges, register the engine via CalculationOptions, evaluate the formula, and save the workbook.
    public class MyCustomEngine : AbstractCalculationEngine
    {
        // Example: force recalculation for a volatile custom function
        public override bool ForceRecalculate(string functionName)
        {
            return functionName.Equals("MYFUNC", StringComparison.OrdinalIgnoreCase);
        }

        // Core calculation logic for custom functions
        public override void Calculate(CalculationData data)
        {
            // Handle only the custom function we are interested in
            if (data.FunctionName.Equals("MYFUNC", StringComparison.OrdinalIgnoreCase))
            {
                // Expecting two parameters; retrieve their values
                if (data.ParamCount >= 2)
                {
                    // Parameters are returned as ReferredArea objects when they refer to cell ranges
                    var param0 = (ReferredArea)data.GetParamValue(0);
                    var param1 = (ReferredArea)data.GetParamValue(1);

                    // Extract the first cell value from each area (simple example)
                    double val0 = Convert.ToDouble(param0.GetValue(0, 0));
                    double val1 = Convert.ToDouble(param1.GetValue(0, 0));

                    // Set the calculated result (sum in this case)
                    data.CalculatedValue = val0 + val1;
                }
                else
                {
                    // Not enough parameters – return an error value
                    data.CalculatedValue = "#VALUE!";
                }
            }
            // For all other functions, let the default engine handle them
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: use provided creation method)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);

            // Set a formula that uses the custom function "MYFUNC"
            sheet.Cells["A3"].Formula = "=MYFUNC(A1, A2)";

            // Configure calculation options to use the custom engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new MyCustomEngine(),
                IgnoreError = false,
                Recursive = true
            };

            // Perform calculation with the custom engine
            workbook.CalculateFormula(options);

            // Output the result of the custom function
            Console.WriteLine("Result of MYFUNC(A1, A2): " + sheet.Cells["A3"].Value);

            // Save the workbook (lifecycle rule: use provided save method)
            workbook.Save("CustomEngineResult.xlsx");
        }
    }
}
