using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace CustomCalculationEngineDemo
{
    // Custom engine that provides implementations for user‑defined functions.
    // If a function is not recognized, it returns the Excel error "#NAME?".
    public class UserDefinedEngine : AbstractCalculationEngine
    {
        // Mapping of function name (upper case) to a delegate that performs the calculation.
        private readonly Dictionary<string, Func<CalculationData, object>> _functions;

        public UserDefinedEngine()
        {
            _functions = new Dictionary<string, Func<CalculationData, object>>(StringComparer.OrdinalIgnoreCase)
            {
                { "MYADD",  CalculateMyAdd },
                { "MYMULT", CalculateMyMult }
                // Add more custom functions here as needed.
            };
        }

        // Force recalculation for all custom functions to ensure each cell gets its own result.
        public override bool ForceRecalculate(string functionName)
        {
            return _functions.ContainsKey(functionName);
        }

        // Core calculation method invoked by Aspose.Cells for each function call.
        public override void Calculate(CalculationData data)
        {
            if (data == null) return;

            // Try to find a user‑defined implementation.
            if (_functions.TryGetValue(data.FunctionName, out var handler))
            {
                // Execute the custom logic and assign the result.
                data.CalculatedValue = handler(data);
                return;
            }

            // Function not found – return Excel's #NAME? error.
            data.CalculatedValue = "#NAME?";
        }

        // Example implementation: MYADD(param1, param2, ...) returns the sum of all numeric parameters.
        private object CalculateMyAdd(CalculationData data)
        {
            double sum = 0;
            for (int i = 0; i < data.ParamCount; i++)
            {
                // Parameters are returned as objects; attempt conversion to double.
                object val = data.GetParamValue(i);
                if (val is double d)
                {
                    sum += d;
                }
                else
                {
                    // Try to convert other types (e.g., string representations of numbers).
                    if (double.TryParse(Convert.ToString(val), out double parsed))
                        sum += parsed;
                }
            }
            return sum;
        }

        // Example implementation: MYMULT(param1, param2, ...) returns the product of all numeric parameters.
        private object CalculateMyMult(CalculationData data)
        {
            double product = 1;
            for (int i = 0; i < data.ParamCount; i++)
            {
                object val = data.GetParamValue(i);
                if (val is double d)
                {
                    product *= d;
                }
                else
                {
                    if (double.TryParse(Convert.ToString(val), out double parsed))
                        product *= parsed;
                }
            }
            return product;
        }
    }

    class Program
    {
        static void Main()
        {
            // -------------------- Create workbook --------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data.
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["A2"].PutValue(10);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["B2"].PutValue(3);

            // Formulas that use user‑defined functions.
            sheet.Cells["C1"].Formula = "=MYADD(A1, A2)";      // Expected 15
            sheet.Cells["C2"].Formula = "=MYMULT(B1, B2)";    // Expected 6
            // This function does not exist in our engine – should return #NAME?.
            sheet.Cells["C3"].Formula = "=UNKNOWNFUNC(A1)";

            // -------------------- Set calculation options with custom engine --------------------
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new UserDefinedEngine(),
                // Keep default settings for other options.
                Recursive = true,
                IgnoreError = false
            };

            // Perform calculation using the custom engine.
            workbook.CalculateFormula(options);

            // Output results to console.
            Console.WriteLine("C1 (MYADD)   = " + sheet.Cells["C1"].Value);
            Console.WriteLine("C2 (MYMULT)  = " + sheet.Cells["C2"].Value);
            Console.WriteLine("C3 (UNKNOWN) = " + sheet.Cells["C3"].Value);

            // -------------------- Save workbook --------------------
            workbook.Save("CustomEngineResult.xlsx");
        }
    }
}