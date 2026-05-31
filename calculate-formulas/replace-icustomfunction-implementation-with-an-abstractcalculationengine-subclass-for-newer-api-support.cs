using System;
using Aspose.Cells;

namespace AsposeCellsCustomEngineDemo
{
    // Custom calculation engine that replaces the old ICustomFunction implementation
    public class MyCustomEngine : AbstractCalculationEngine
    {
        // Example: indicate that we want to force recalculation for a volatile custom function
        public override bool ForceRecalculate(string functionName)
        {
            return functionName.Equals("MYVOLATILEFUNC", StringComparison.OrdinalIgnoreCase);
        }

        // Example: we do not need array mode or literal text for parameters
        public override bool IsParamArrayModeRequired => false;
        public override bool IsParamLiteralRequired => false;

        // Core calculation logic for custom functions
        public override void Calculate(CalculationData data)
        {
            // Handle a custom function named MYFUNC that adds two numbers
            if (data.FunctionName.Equals("MYFUNC", StringComparison.OrdinalIgnoreCase))
            {
                // Retrieve parameter values; they are returned as ReferredArea objects
                var param0 = (ReferredArea)data.GetParamValue(0);
                var param1 = (ReferredArea)data.GetParamValue(1);

                double val0 = Convert.ToDouble(param0.GetValue(0, 0));
                double val1 = Convert.ToDouble(param1.GetValue(0, 0));

                // Set the result back to the engine
                data.CalculatedValue = val0 + val1;
                return;
            }

            // Handle a custom volatile function that returns the current timestamp
            if (data.FunctionName.Equals("MYVOLATILEFUNC", StringComparison.OrdinalIgnoreCase))
            {
                data.CalculatedValue = DateTime.Now.ToString("O");
                return;
            }

            // For any other function, let the default engine handle it
            // (do not set CalculatedValue, so Aspose.Cells will calculate normally)
        }
    }

    // Optional: custom function definition to force array mode for a parameter (if needed)
    public class MyCustomFunctionDefinition : CustomFunctionDefinition
    {
        // Example: make the first parameter of MYARRAYFUNC be calculated in array mode
        public override int[] GetArrayModeParameters(string functionName)
        {
            if (functionName.Equals("MYARRAYFUNC", StringComparison.OrdinalIgnoreCase))
                return new[] { 0 }; // first parameter
            return base.GetArrayModeParameters(functionName);
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells used by custom functions
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["B1"].Formula = "=MYFUNC(A1, A2)";          // uses MyCustomEngine
            sheet.Cells["B2"].Formula = "=MYVOLATILEFUNC()";       // volatile example

            // Set up calculation options with the custom engine
            CalculationOptions calcOptions = new CalculationOptions
            {
                CustomEngine = new MyCustomEngine(),
                IgnoreError = false,
                Recursive = true
            };

            // If we need special parameter handling, update the custom function definition
            workbook.UpdateCustomFunctionDefinition(new MyCustomFunctionDefinition());

            // Perform calculation using the custom engine
            workbook.CalculateFormula(calcOptions);

            // Output results
            Console.WriteLine("Result of MYFUNC(A1, A2): " + sheet.Cells["B1"].Value);
            Console.WriteLine("Result of MYVOLATILEFUNC(): " + sheet.Cells["B2"].Value);

            // Save the workbook (demonstrates lifecycle rule usage)
            workbook.Save("CustomEngineResult.xlsx");
        }
    }
}