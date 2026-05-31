using System;
using Aspose.Cells;

namespace CustomCalculationEngineDemo
{
    // Custom engine that handles a user‑defined function MYFUNC
    public class MyCustomEngine : AbstractCalculationEngine
    {
        // Override Calculate to provide custom logic
        public override void Calculate(CalculationData data)
        {
            // Check if the function being calculated is MYFUNC
            if (data.FunctionName.Equals("MYFUNC", StringComparison.OrdinalIgnoreCase))
            {
                // Retrieve the two parameters passed to MYFUNC
                object param0 = data.GetParamValue(0);
                object param1 = data.GetParamValue(1);

                // Convert parameters to double and compute the sum
                double val0 = Convert.ToDouble(param0);
                double val1 = Convert.ToDouble(param1);
                double result = val0 + val1;

                // Set the calculated result so Aspose.Cells can use it
                data.CalculatedValue = result;
            }
        }

        // No special handling for shared formulas; use default behavior
        public override bool ForceRecalculate(string functionName) => false;
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells that will be used as arguments for the custom function
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);

            // Set a formula that calls the custom function MYFUNC(A1, A2)
            sheet.Cells["B1"].Formula = "=MYFUNC(A1, A2)";

            // Create calculation options and assign the custom engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new MyCustomEngine()
            };

            // Perform calculation using the options (custom engine will be invoked)
            workbook.CalculateFormula(options);

            // Output the result of the custom function
            Console.WriteLine("Result of MYFUNC(A1, A2): " + sheet.Cells["B1"].Value);

            // Save the workbook (uses the standard save rule)
            workbook.Save("CustomEngineResult.xlsx");
        }
    }
}