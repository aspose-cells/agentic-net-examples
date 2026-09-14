// Title: Create a custom calculation engine in Aspose.Cells for .NET to evaluate a MYADD user-defined function
// AI Prompts: Subclass AbstractCalculationEngine and override the Calculate method to recognize the MYADD function, convert its two parameters to numbers, and return their sum. | Configure CalculationOptions.CustomEngine with the new engine, place =MYADD formulas in worksheet cells, invoke workbook.CalculateFormula, and save the workbook with the computed results.
// Common Searches: how to implement a user defined function in Aspose.Cells using AbstractCalculationEngine | Aspose.Cells .NET custom calculation engine example with MYADD function | set CustomEngine in CalculationOptions to calculate custom formulas in a workbook | force recalculation for a specific user defined function in Aspose.Cells
// Tags: Aspose.Cells custom calculation engine .NET | user defined function MYADD implementation | override AbstractCalculationEngine Calculate method | assign CustomEngine in CalculationOptions | force recalculate custom formula Aspose.Cells

using System;
using Aspose.Cells;

namespace CustomEngineDemo
{
    // Custom calculation engine that implements a simple function MYADD which adds two numbers
    // The example shows how to derive a MyAddEngine from AbstractCalculationEngine, override Calculate to process the MYADD function by adding two numeric arguments, force recalculation for that function, assign the engine via CalculationOptions.CustomEngine, evaluate cells containing =MYADD formulas, output the results, and save the workbook.
    public class MyAddEngine : AbstractCalculationEngine
    {
        // Override the Calculate method to provide custom logic
        public override void Calculate(CalculationData data)
        {
            // Check if the function being calculated is our custom function
            if (data.FunctionName.Equals("MYADD", StringComparison.OrdinalIgnoreCase))
            {
                // Ensure we have exactly two parameters
                if (data.ParamCount == 2)
                {
                    // Retrieve the parameter values; they are returned as objects
                    object param0 = data.GetParamValue(0);
                    object param1 = data.GetParamValue(1);

                    // Convert parameters to double (handles numeric cells and literals)
                    double val0 = Convert.ToDouble(param0);
                    double val1 = Convert.ToDouble(param1);

                    // Set the calculated result
                    data.CalculatedValue = val0 + val1;
                }
                else
                {
                    // If parameter count is unexpected, return an error value
                    data.CalculatedValue = "#VALUE!";
                }
            }
        }

        // Force recalculation for the custom function so that each cell gets its own result
        public override bool ForceRecalculate(string functionName)
        {
            return functionName.Equals("MYADD", StringComparison.OrdinalIgnoreCase);
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["B1"].PutValue(5);
            sheet.Cells["B2"].PutValue(15);

            // Use the custom function MYADD in two different cells
            sheet.Cells["C1"].Formula = "=MYADD(A1, A2)"; // Expected result: 30
            sheet.Cells["C2"].Formula = "=MYADD(B1, B2)"; // Expected result: 20

            // Set calculation options with the custom engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new MyAddEngine()
            };

            // Perform calculation using the custom engine
            workbook.CalculateFormula(options);

            // Output the results to the console
            Console.WriteLine("C1 (MYADD A1,A2) = " + sheet.Cells["C1"].Value);
            Console.WriteLine("C2 (MYADD B1,B2) = " + sheet.Cells["C2"].Value);

            // Save the workbook (the file will contain the calculated values)
            workbook.Save("CustomEngineDemo.xlsx");
        }
    }
}
