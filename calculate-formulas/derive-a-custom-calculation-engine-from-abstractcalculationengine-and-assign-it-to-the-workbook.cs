using System;
using Aspose.Cells;

namespace CustomEngineDemo
{
    // Custom calculation engine derived from AbstractCalculationEngine
    public class DoubleEngine : AbstractCalculationEngine
    {
        // Override Calculate to handle custom function "DOUBLE"
        public override void Calculate(CalculationData data)
        {
            // Check if the function name matches our custom function (case‑insensitive)
            if (string.Equals(data.FunctionName, "DOUBLE", StringComparison.OrdinalIgnoreCase))
            {
                // Expect a single numeric parameter
                if (data.ParamCount == 1)
                {
                    // Retrieve the parameter value; it may be a ReferredArea or a direct value
                    object param = data.GetParamValue(0);
                    double number;

                    // If the parameter is a ReferredArea, get the first cell's value
                    if (param is ReferredArea area)
                    {
                        number = Convert.ToDouble(area.GetValue(0, 0));
                    }
                    else
                    {
                        number = Convert.ToDouble(param);
                    }

                    // Set the calculated result (double the input)
                    data.CalculatedValue = number * 2;
                }
                else
                {
                    // Incorrect number of arguments – return Excel error
                    data.CalculatedValue = "#VALUE!";
                }
            }
            // For all other functions, do nothing and let the default engine handle them
        }

        // Ensure the engine recalculates the function for each cell (optional)
        public override bool ForceRecalculate(string functionName)
        {
            return string.Equals(functionName, "DOUBLE", StringComparison.OrdinalIgnoreCase);
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];

            // Put a sample value in A1
            sheet.Cells["A1"].PutValue(5);

            // Use the custom function in B1
            sheet.Cells["B1"].Formula = "=DOUBLE(A1)";

            // Configure calculation options with the custom engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new DoubleEngine()
            };

            // Calculate formulas using the custom engine
            wb.CalculateFormula(options);

            // Output the result to console
            Console.WriteLine("Result of DOUBLE(A1): " + sheet.Cells["B1"].Value);

            // Save the workbook (Excel and PDF formats)
            wb.Save("CustomEngineResult.xlsx");
            wb.Save("CustomEngineResult.pdf");
        }
    }
}