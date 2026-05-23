using System;
using Aspose.Cells;

namespace AsposeCellsCustomEngineDemo
{
    // Custom calculation engine that implements a user‑defined function MYFUNC
    public class MyCustomEngine : AbstractCalculationEngine
    {
        // This method is called for each custom function encountered during calculation
        public override void Calculate(CalculationData data)
        {
            // Check for the custom function name (case‑insensitive)
            if (data.FunctionName.Equals("MYFUNC", StringComparison.OrdinalIgnoreCase))
            {
                // Expect exactly two parameters
                if (data.ParamCount == 2)
                {
                    // Retrieve the first parameter value
                    object param1 = data.GetParamValue(0);
                    // Retrieve the second parameter value
                    object param2 = data.GetParamValue(1);

                    // Convert parameters to double (they may be numeric or ReferredArea)
                    double val1 = ConvertToDouble(param1);
                    double val2 = ConvertToDouble(param2);

                    // Example logic: return the product of the two parameters
                    data.CalculatedValue = val1 * val2;
                }
                else
                {
                    // Incorrect number of arguments – return a #VALUE! error
                    data.CalculatedValue = "#VALUE!";
                }
            }
            // For all other functions the default engine will be used automatically
        }

        // Helper to extract a double from supported parameter types
        private double ConvertToDouble(object param)
        {
            if (param is double d)
                return d;

            if (param is ReferredArea ra)
                return Convert.ToDouble(ra.GetValue(0, 0));

            // Fallback conversion
            return Convert.ToDouble(param);
        }
    }

    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a new workbook and add sample data
            // -------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Simple numeric values
            cells["A1"].PutValue(5);
            cells["A2"].PutValue(10);

            // Formula that uses the custom function MYFUNC
            cells["A3"].Formula = "=MYFUNC(A1, A2)";

            // Formula that uses a built‑in function (to show default handling)
            cells["B1"].Formula = "=SUM(A1:A2)";

            // -------------------------------------------------
            // 2. Configure calculation options with the custom engine
            // -------------------------------------------------
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new MyCustomEngine(),
                // Optional: ignore errors during calculation
                IgnoreError = true,
                // Optional: ensure dependent cells are calculated recursively
                Recursive = true
            };

            // -------------------------------------------------
            // 3. Evaluate all formulas in the workbook using the custom engine
            // -------------------------------------------------
            workbook.CalculateFormula(options);

            // -------------------------------------------------
            // 4. Output the results to the console
            // -------------------------------------------------
            Console.WriteLine("Result of MYFUNC(A1, A2) in A3: " + cells["A3"].Value);
            Console.WriteLine("Result of SUM(A1:A2) in B1: " + cells["B1"].Value);

            // -------------------------------------------------
            // 5. Save the workbook (optional)
            // -------------------------------------------------
            workbook.Save("CustomEngineResult.xlsx");
        }
    }
}