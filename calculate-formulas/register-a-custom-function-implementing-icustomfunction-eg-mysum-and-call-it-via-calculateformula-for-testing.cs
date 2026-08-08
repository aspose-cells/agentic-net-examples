// Title: Aspose.Cells C# – Register a Custom ICustomFunction (MYSUM) and Run it with CalculateFormula
// Description: Learn how to implement a user‑defined function by extending AbstractCalculationEngine, register it as "MYSUM", apply the formula =MYSUM(A1,A2,A3) in a workbook, and evaluate the result with CalculationOptions.CustomEngine in C#.
// Keywords: Aspose.Cells custom function C# | ICustomFunction implementation | CalculateFormula custom engine | MYSUM user defined function | ReferredArea handling | .NET spreadsheet calculation
// Common Searches: custom ICustomFunction Aspose.Cells example | how to add user defined function in Aspose.Cells .NET | calculate formula with custom engine C# | sum cells using MYSUM Aspose.Cells | register and call custom function in workbook
// Developer Intent: Create and register a custom ICustomFunction named MYSUM, then evaluate it through CalculateFormula in a C# Aspose.Cells workbook.
// Use Cases: Extend spreadsheet calculations with proprietary functions such as MYSUM. | Process cell references and ranges inside a custom engine using ReferredArea objects. | Integrate custom formula evaluation into automated .NET reporting pipelines.
// AI Prompts: Write C# code that defines a custom ICustomFunction called MYSUM for Aspose.Cells and uses CalculateFormula to sum cells A1 to A3. | Show how to handle ReferredArea parameters in an AbstractCalculationEngine implementation. | Explain the steps to configure CalculationOptions.CustomEngine and retrieve the result of a user‑defined function.

using System;
using Aspose.Cells;

namespace AsposeCellsCustomFunctionDemo
{
    // Custom calculation engine that implements a user‑defined function "MYSUM"
    // Learn how to implement a user‑defined function by extending AbstractCalculationEngine, register it as "MYSUM", apply the formula =MYSUM(A1,A2,A3) in a workbook, and evaluate the result with CalculationOptions.CustomEngine in C#.
    class MyCustomEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // Handle only the custom function we are interested in
            if (data.FunctionName.Equals("MYSUM", StringComparison.OrdinalIgnoreCase))
            {
                double sum = 0;

                // Iterate over all parameters passed to the function
                for (int i = 0; i < data.ParamCount; i++)
                {
                    // Parameters are returned as objects (numbers, strings, ReferredArea, etc.)
                    object param = data.GetParamValue(i);

                    // If the parameter is a ReferredArea (e.g., a cell reference), extract its value
                    if (param is ReferredArea area)
                    {
                        // For a single‑cell reference GetValue(0,0) returns the cell's value
                        sum += Convert.ToDouble(area.GetValue(0, 0));
                    }
                    else
                    {
                        // Direct numeric or convertible value
                        sum += Convert.ToDouble(param);
                    }
                }

                // Set the result that will be written back to the worksheet cell
                data.CalculatedValue = sum;
            }
        }

        // No special force‑recalculation logic needed for this demo
        public override bool ForceRecalculate(string functionName) => false;
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data that will be used as arguments for the custom function
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["A2"].PutValue(7);
            sheet.Cells["A3"].PutValue(12);

            // Set a formula that calls the custom function "MYSUM"
            // The function will sum the values of A1, A2 and A3
            sheet.Cells["B1"].Formula = "=MYSUM(A1, A2, A3)";

            // Configure calculation options to use the custom engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new MyCustomEngine()
            };

            // Perform calculation using the custom engine (lifecycle rule: calculate)
            workbook.CalculateFormula(options);

            // Output the result of the custom function
            Console.WriteLine("Result of MYSUM(A1, A2, A3): " + sheet.Cells["B1"].Value);

            // Save the workbook (lifecycle rule: save)
            workbook.Save("CustomFunctionResult.xlsx");
        }
    }
}
