using System;
using Aspose.Cells;

namespace CustomEngineDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook in XLSX format
            Workbook wb = new Workbook("input.xlsx");

            // Create a custom calculation engine instance
            var engine = new MyCustomEngine();

            // Set calculation options to use the custom engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = engine,
                IgnoreError = false,
                Recursive = true
            };

            // If the custom function requires special parameter handling, update its definition
            wb.UpdateCustomFunctionDefinition(new MyCustomFunctionDefinition());

            // Calculate all formulas in the workbook using the custom engine
            wb.CalculateFormula(options);

            // Save the workbook after calculation
            wb.Save("output.xlsx");
        }
    }

    // Custom calculation engine that extends the default engine
    public class MyCustomEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // Handle a custom function named "MYFUNC"
            if (string.Equals(data.FunctionName, "MYFUNC", StringComparison.OrdinalIgnoreCase))
            {
                // Expect at least two parameters
                if (data.ParamCount >= 2)
                {
                    // Retrieve the first two parameter values (they may be ReferredArea or direct values)
                    double val1 = Convert.ToDouble(GetNumericValue(data.GetParamValue(0)));
                    double val2 = Convert.ToDouble(GetNumericValue(data.GetParamValue(1)));

                    // Simple custom logic: sum the two parameters
                    data.CalculatedValue = val1 + val2;
                }
                else
                {
                    // Not enough parameters – return an error value
                    data.CalculatedValue = "#VALUE!";
                }
            }
        }

        // Helper method to extract a numeric value from a parameter that could be a ReferredArea
        private object GetNumericValue(object param)
        {
            if (param is ReferredArea area)
            {
                // Return the value of the first cell in the referred area
                return area.GetValue(0, 0);
            }
            return param;
        }

        // Force recalculation for the custom function on each calculation pass
        public override bool ForceRecalculate(string functionName)
        {
            return string.Equals(functionName, "MYFUNC", StringComparison.OrdinalIgnoreCase);
        }
    }

    // Custom function definition (optional) – can specify which parameters need array‑mode evaluation
    public class MyCustomFunctionDefinition : CustomFunctionDefinition
    {
        public override int[] GetArrayModeParameters(string functionName)
        {
            if (string.Equals(functionName, "MYFUNC", StringComparison.OrdinalIgnoreCase))
            {
                // No parameters require array mode for this example
                return new int[0];
            }
            return base.GetArrayModeParameters(functionName);
        }
    }
}