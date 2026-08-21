// Title: C# – Assign a Custom Calculation Engine in Aspose.Cells to Implement a DOUBLE Function
// Description: Shows how to inherit from AbstractCalculationEngine, override Calculate to support a user‑defined DOUBLE function, attach the engine via CalculationOptions.CustomEngine, run wb.CalculateFormula, and save the workbook.
// Keywords: Aspose.Cells | CustomCalculationEngine | AbstractCalculationEngine | C# user defined function | DOUBLE UDF | CalculationOptions.CustomEngine | Workbook.CalculateFormula | ReferredArea handling | Excel formula extension | UDF debugging
// Common Searches: Aspose.Cells custom calculation engine example | Create user defined function in Aspose.Cells C# | Set CustomEngine for workbook Aspose.Cells | Override Calculate method AbstractCalculationEngine | DOUBLE function Aspose.Cells
// Developer Intent: Build a custom calculation engine that evaluates a DOUBLE user‑defined function and assign it to a workbook for formula calculation.
// Use Cases: Double numeric values in generated reports using the =DOUBLE() formula. | Extend Aspose.Cells with multiple proprietary functions while preserving built‑in calculations. | Provide deterministic results for custom financial or scientific formulas in automated spreadsheet processing.
// AI Prompts: Write C# code for a custom calculation engine that adds a TRIPLE function returning three times the input. | Explain how to register several user‑defined functions in a single AbstractCalculationEngine and apply it to a workbook. | Describe steps to test and debug a custom calculation engine, including handling ReferredArea parameters and returning error values.

using System;
using Aspose.Cells;

namespace CustomEngineDemo
{
    // Custom calculation engine derived from AbstractCalculationEngine
    // Shows how to inherit from AbstractCalculationEngine, override Calculate to support a user‑defined DOUBLE function, attach the engine via CalculationOptions.CustomEngine, run wb.CalculateFormula, and save the workbook.
    public class DoubleEngine : AbstractCalculationEngine
    {
        // Override Calculate to handle custom function "DOUBLE"
        public override void Calculate(CalculationData data)
        {
            // Check if the function name matches our custom function (case‑insensitive)
            if (string.Equals(data.FunctionName, "DOUBLE", StringComparison.OrdinalIgnoreCase))
            {
                // Expect at least one parameter
                if (data.ParamCount > 0)
                {
                    // Get the first parameter value (it may be a ReferredArea, double, etc.)
                    object param = data.GetParamValue(0);

                    double value;

                    // If the parameter is a ReferredArea (range), take the first cell value
                    if (param is ReferredArea area)
                    {
                        value = Convert.ToDouble(area.GetValue(0, 0));
                    }
                    else
                    {
                        // Direct value conversion
                        value = Convert.ToDouble(param);
                    }

                    // Set the calculated result (double the input)
                    data.CalculatedValue = value * 2;
                }
                else
                {
                    // No parameters – return an error value
                    data.CalculatedValue = "#VALUE!";
                }
            }
            // For all other functions, do nothing and let the default engine handle them
        }

        // Optional: force recalculation for volatile custom functions
        public override bool ForceRecalculate(string functionName)
        {
            // Our function is deterministic, so no forced recalculation needed
            return false;
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook (using the create rule)
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];

            // Put sample data in A1
            sheet.Cells["A1"].PutValue(7);

            // Use the custom function "DOUBLE" in cell B1
            sheet.Cells["B1"].Formula = "=DOUBLE(A1)";

            // Set calculation options with the custom engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new DoubleEngine()
            };

            // Calculate formulas using the custom engine
            wb.CalculateFormula(options);

            // Output the result to console
            Console.WriteLine("Result of DOUBLE(A1): " + sheet.Cells["B1"].Value);

            // Save the workbook (using the save rule)
            wb.Save("CustomEngineDemo.xlsx");
        }
    }
}
