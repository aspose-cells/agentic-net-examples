using System;
using Aspose.Cells;

namespace CustomEngineDemo
{
    // Custom calculation engine that implements a simple DOUBLE function
    public class DoubleEngine : AbstractCalculationEngine
    {
        // Override the Calculate method to handle custom functions
        public override void Calculate(CalculationData data)
        {
            // Check if the function being calculated is our custom "DOUBLE"
            if (data.FunctionName.Equals("DOUBLE", StringComparison.OrdinalIgnoreCase))
            {
                // Retrieve the first parameter value
                object param = data.GetParamValue(0);

                double number;

                // The parameter may be a direct numeric value or a ReferredArea (range)
                if (param is ReferredArea area)
                {
                    // Get the value from the first cell of the range
                    number = Convert.ToDouble(area.GetValue(0, 0));
                }
                else
                {
                    // Assume it's a numeric value
                    number = Convert.ToDouble(param);
                }

                // Set the calculated result (double the input)
                data.CalculatedValue = number * 2;
            }
        }

        // Ensure the custom function is recalculated for each cell (optional)
        public override bool ForceRecalculate(string functionName)
        {
            return functionName.Equals("DOUBLE", StringComparison.OrdinalIgnoreCase);
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a sample value in A1
            sheet.Cells["A1"].PutValue(5);

            // Use the custom function in cell B1
            sheet.Cells["B1"].Formula = "=DOUBLE(A1)";

            // Set calculation options with the custom engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new DoubleEngine()
            };

            // Calculate formulas using the custom engine
            workbook.CalculateFormula(options);

            // Output the result
            Console.WriteLine("Result of DOUBLE(A1): " + sheet.Cells["B1"].Value);

            // Save the workbook
            workbook.Save("CustomEngineDemo.xlsx");
        }
    }
}