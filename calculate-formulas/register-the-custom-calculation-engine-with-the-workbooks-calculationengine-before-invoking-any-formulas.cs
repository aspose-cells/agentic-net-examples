// Title: How to register a custom calculation engine in Aspose.Cells (C#) to evaluate a user‑defined function
// AI Prompts: Create a class inheriting from AbstractCalculationEngine that implements MYFUNC logic and assign it to CalculationOptions.CustomEngine. | Configure CalculationOptions with the custom engine and call Workbook.CalculateFormula to compute cells containing the MYFUNC formula. | Write the workbook to an .xlsx file after the custom engine calculation to persist the computed result.
// Common Searches: Aspose.Cells C# example for adding a user‑defined function with a custom engine | How to use CalculationOptions.CustomEngine to run custom formulas in Aspose.Cells | Registering and invoking a custom AbstractCalculationEngine before workbook.CalculateFormula in .NET
// Tags: custom calculation engine registration Aspose.Cells C# | user-defined function AbstractCalculationEngine | assign CalculationOptions.CustomEngine workbook | formula evaluation using custom calculation engine | persist custom function result Excel file

using System;
using Aspose.Cells;

namespace CustomCalculationEngineDemo
{
    // Custom engine that implements a user‑defined function MYFUNC
    // The sample defines MyCustomEngine derived from AbstractCalculationEngine to handle a custom function MYFUNC, registers it via CalculationOptions.CustomEngine, calculates the formula '=MYFUNC(A1, A2)' with Workbook.CalculateFormula, outputs the result, and saves the workbook as an .xlsx file.
    public class MyCustomEngine : AbstractCalculationEngine
    {
        // Override Calculate to handle the custom function
        public override void Calculate(CalculationData data)
        {
            // Check if the function being evaluated is MYFUNC (case‑insensitive)
            if (string.Equals(data.FunctionName, "MYFUNC", StringComparison.OrdinalIgnoreCase))
            {
                // Retrieve the first two parameters passed to the function
                object param0 = data.GetParamValue(0);
                object param1 = data.GetParamValue(1);

                // Convert parameters to double (they may be numeric or ReferredArea objects)
                double val0 = Convert.ToDouble(param0);
                double val1 = Convert.ToDouble(param1);

                // Simple custom logic: return the sum of the two parameters multiplied by 10
                data.CalculatedValue = (val0 + val1) * 10;
            }
        }

        // No special handling for shared formulas; default behavior is sufficient
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
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["A2"].PutValue(7);

            // Set a formula that calls the custom function MYFUNC
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

            // Save the workbook (optional, demonstrates that the result is persisted)
            workbook.Save("CustomEngineResult.xlsx");
        }
    }
}
