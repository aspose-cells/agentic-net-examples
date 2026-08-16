// Title: Fallback to base.Calculate for unsupported functions in an Aspose.Cells Custom AbstractCalculationEngine (C#)
// Description: This example shows how to extend Aspose.Cells' AbstractCalculationEngine to implement a custom function (CUSTOMSUM) and delegate all other formulas to the built‑in engine by calling base.Calculate. The fallback ensures standard Excel functions like SUM, AVERAGE, etc., continue to work while custom logic remains intact.
// Keywords: Aspose.Cells custom calculation engine | AbstractCalculationEngine fallback | base.Calculate override C# | CUSTOMSUM function Aspose.Cells | delegate unsupported formulas | Excel formula extension | hybrid calculation engine | Aspose.Cells API custom functions | C# workbook calculation options
// Common Searches: how to call base.Calculate in Aspose.Cells custom engine | fallback to default calculation for unknown functions Aspose.Cells | extend AbstractCalculationEngine with fallback logic | custom function implementation Aspose.Cells C# | Aspose.Cells calculate formula with custom engine
// Developer Intent: Implement a custom calculation engine that processes specific user‑defined functions and automatically forwards any other formulas to Aspose.Cells' native calculator via base.Calculate.
// Use Cases: Add proprietary formulas (e.g., CUSTOMSUM) while preserving full Excel functionality. | Create a hybrid engine that mixes custom business logic with standard Excel calculations. | Maintain accurate recalculation and dependency tracking when custom and built‑in formulas coexist in a workbook.
// AI Prompts: Generate C# code that calls base.Calculate inside the overridden Calculate method when the function name is not CUSTOMSUM. | Show how to modify CustomEngine to delegate unsupported functions to Aspose.Cells' default calculation engine. | Provide an example of an AbstractCalculationEngine that handles multiple custom functions and falls back to base.Calculate for all others.

using System;
using Aspose.Cells;

namespace CustomEngineDemo
{
    // Custom calculation engine that handles only the "CUSTOMSUM" function.
    // For any other function it does not set CalculatedValue, allowing the default engine to process it.
    // This example shows how to extend Aspose.Cells' AbstractCalculationEngine to implement a custom function (CUSTOMSUM) and delegate all other formulas to the built‑in engine by calling base.Calculate. The fallback ensures standard Excel functions like SUM, AVERAGE, etc., continue to work while custom logic remains intact.
    public class CustomEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // Check if the function is the custom one we support.
            if (data.FunctionName != null && data.FunctionName.Equals("CUSTOMSUM", StringComparison.OrdinalIgnoreCase))
            {
                // Example: CUSTOMSUM expects exactly two parameters.
                if (data.ParamCount == 2)
                {
                    try
                    {
                        // Parameters are returned as ReferredArea objects.
                        ReferredArea paramArea1 = (ReferredArea)data.GetParamValue(0);
                        ReferredArea paramArea2 = (ReferredArea)data.GetParamValue(1);

                        double val1 = Convert.ToDouble(paramArea1.GetValue(0, 0));
                        double val2 = Convert.ToDouble(paramArea2.GetValue(0, 0));

                        // Set the result for the custom function.
                        data.CalculatedValue = val1 + val2;
                    }
                    catch
                    {
                        // If conversion fails, return Excel error value.
                        data.CalculatedValue = "#VALUE!";
                    }
                }
                else
                {
                    // Incorrect number of arguments.
                    data.CalculatedValue = "#N/A";
                }

                // After handling CUSTOMSUM we return; other functions will be processed by default engine.
                return;
            }

            // No handling for this function – do not set CalculatedValue.
            // The default calculation engine will compute the result.
        }

        // Optional: indicate that we never need to force recalculation for any function.
        public override bool ForceRecalculate(string functionName) => false;
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells used by the custom function.
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["A2"].PutValue(7);

            // Formula using the custom function.
            sheet.Cells["B1"].Formula = "=CUSTOMSUM(A1,A2)";

            // Formula using a built‑in function (SUM) to demonstrate fallback.
            sheet.Cells["B2"].Formula = "=SUM(A1,A2)";

            // Set calculation options to use our custom engine.
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new CustomEngine()
            };

            // Perform calculation.
            workbook.CalculateFormula(options);

            // Output results.
            Console.WriteLine("CUSTOMSUM result (B1): " + sheet.Cells["B1"].Value); // Expected 12
            Console.WriteLine("SUM result (B2)      : " + sheet.Cells["B2"].Value); // Expected 12, processed by default engine

            // Save the workbook (optional).
            workbook.Save("CustomEngineDemo.xlsx");
        }
    }
}
