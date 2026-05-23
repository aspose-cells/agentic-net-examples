using System;
using Aspose.Cells;

namespace CustomEngineDemo
{
    // Custom calculation engine that handles a specific function and falls back to the default engine for others
    public class MyCustomEngine : AbstractCalculationEngine
    {
        // Override Calculate to process only the custom function "MYFUNC"
        public override void Calculate(CalculationData data)
        {
            // Check if the function name matches our custom function (case‑insensitive)
            if (data.FunctionName != null && data.FunctionName.Equals("MYFUNC", StringComparison.OrdinalIgnoreCase))
            {
                // Example: MYFUNC takes two numeric parameters and returns their sum
                try
                {
                    // Retrieve the first parameter value
                    object param0 = data.GetParamValue(0);
                    // Retrieve the second parameter value
                    object param1 = data.GetParamValue(1);

                    // Convert parameters to double (handles numbers, strings that can be parsed, etc.)
                    double val0 = Convert.ToDouble(param0);
                    double val1 = Convert.ToDouble(param1);

                    // Set the calculated result – this overrides the default calculation for MYFUNC
                    data.CalculatedValue = val0 + val1;
                }
                catch
                {
                    // If conversion fails, return Excel error #VALUE!
                    data.CalculatedValue = "#VALUE!";
                }
            }
            // For any other function we do NOT set CalculatedValue.
            // Leaving it unset tells Aspose.Cells to use its built‑in engine (fallback behavior).
        }

        // No need to force recalculation for this example
        public override bool ForceRecalculate(string functionName) => false;
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells with sample data
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);

            // Use the custom function MYFUNC in a formula
            sheet.Cells["B1"].Formula = "=MYFUNC(A1, A2)";

            // Also add a built‑in function to demonstrate fallback
            sheet.Cells["C1"].Formula = "=SUM(A1, A2)";

            // Configure calculation options to use our custom engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new MyCustomEngine()
            };

            // Perform calculation – MYFUNC will be handled by MyCustomEngine,
            // SUM will be processed by the default engine (fallback)
            workbook.CalculateFormula(options);

            // Output results
            Console.WriteLine("Result of MYFUNC(A1, A2): " + sheet.Cells["B1"].Value); // Expected 30
            Console.WriteLine("Result of SUM(A1, A2): " + sheet.Cells["C1"].Value);   // Expected 30

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("CustomEngineResult.xlsx");
        }
    }
}