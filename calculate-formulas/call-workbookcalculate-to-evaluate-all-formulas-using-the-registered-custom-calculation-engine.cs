// Title: C# – Evaluate All Workbook Formulas with a Custom Calculation Engine in Aspose.Cells
// Description: Demonstrates how to register a user‑defined AbstractCalculationEngine, assign a custom function (MYFUNC) in a formula, call workbook.CalculateFormula with CalculationOptions, and save the computed workbook.
// Keywords: Aspose.Cells custom calculation engine | C# custom function MYFUNC | CalculateFormula with custom engine | AbstractCalculationEngine example | evaluate workbook formulas .NET | Aspose.Cells custom formula implementation | register custom engine Aspose.Cells
// Common Searches: Aspose.Cells custom calculation engine tutorial | How to add user‑defined functions in Aspose.Cells C# | Calculate all formulas with custom engine Aspose.Cells | Register AbstractCalculationEngine for MYFUNC | C# example for workbook.CalculateFormula custom engine
// Developer Intent: Implement and invoke a custom calculation engine to process user‑defined functions across an entire Aspose.Cells workbook.
// Use Cases: Create proprietary business logic (e.g., a custom SUM) that runs during formula evaluation. | Replace the default calculator with a tailored engine to handle special data types or external services. | Generate a spreadsheet, embed custom formulas, compute results programmatically, and persist the final values.
// AI Prompts: Show C# code that defines a class inheriting from AbstractCalculationEngine to handle a function named MYFUNC and uses it with workbook.CalculateFormula. | Provide a step‑by‑step example of registering a custom calculation engine via CalculationOptions in Aspose.Cells. | Explain how to retrieve parameter values inside the Calculate method for both scalar arguments and ReferredArea ranges.

using System;
using Aspose.Cells;

namespace AsposeCellsCustomCalcDemo
{
    // Demonstrates how to register a user‑defined AbstractCalculationEngine, assign a custom function (MYFUNC) in a formula, call workbook.CalculateFormula with CalculationOptions, and save the computed workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and its cells
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);

            // Use a custom function in a formula
            cells["A3"].Formula = "=MYFUNC(A1, A2)";

            // Set up calculation options with a custom engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new MyCustomEngine()
            };

            // Evaluate all formulas using the custom engine
            workbook.CalculateFormula(options);

            // Display the result of the custom function
            Console.WriteLine("Result of MYFUNC(A1, A2): " + cells["A3"].Value);

            // Save the workbook
            workbook.Save("CustomEngineResult.xlsx");
        }
    }

    // Custom calculation engine that implements the MYFUNC function
    public class MyCustomEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // Handle only the custom function MYFUNC
            if (data.FunctionName.Equals("MYFUNC", StringComparison.OrdinalIgnoreCase))
            {
                double sum = 0;

                // Iterate over all parameters passed to the function
                for (int i = 0; i < data.ParamCount; i++)
                {
                    object param = data.GetParamValue(i);

                    // Parameters may be a ReferredArea (cell range) or a direct value
                    if (param is ReferredArea area)
                    {
                        // Get the value of the first cell in the area
                        object val = area.GetValue(0, 0);
                        if (val is double d)
                            sum += d;
                    }
                    else if (param is double d)
                    {
                        sum += d;
                    }
                }

                // Set the calculated result for the function
                data.CalculatedValue = sum;
            }
        }
    }
}
