// Title: Register a custom ICustomFunction (MySum) in Aspose.Cells .NET and evaluate it using Workbook.CalculateFormula
// AI Prompts: Create a class that inherits AbstractCalculationEngine, implement a MYSUM function that adds all numeric arguments (including ReferredArea values), assign CalculationOptions.CustomEngine to this class, set a cell formula =MYSUM(A1,A2), and invoke Workbook.CalculateFormula to obtain the result. | Demonstrate how to expose a user‑defined function to Excel formulas in Aspose.Cells by processing literal and cell‑reference parameters, registering the custom engine, and reading the calculated value from the worksheet.
// Common Searches: asp.net how to add a user defined function to Aspose.Cells calculation engine | example of custom ICustomFunction implementation for summing cells in Aspose.Cells | using CalculationOptions.CustomEngine to run custom formulas in Aspose.Cells .NET | Aspose.Cells calculate custom function MYSUM from worksheet cells | register custom calculation engine for Excel formulas with Aspose.Cells
// Tags: Aspose.Cells custom ICustomFunction | AbstractCalculationEngine sum example | CalculationOptions.CustomEngine usage | Workbook.CalculateFormula with user‑defined function | handling ReferredArea in custom function

using System;
using Aspose.Cells;

namespace AsposeCellsCustomFunctionDemo
{
    // Custom calculation engine that implements a user‑defined function named "MYSUM"
    // The sample defines MySumEngine derived from AbstractCalculationEngine, implements the MYSUM function to sum numeric arguments (including values from ReferredArea cell references), registers this engine via CalculationOptions.CustomEngine, places the formula =MYSUM(A1,A2) in a worksheet cell, calculates the workbook with Workbook.CalculateFormula, prints the result, and saves the file as CustomFunctionDemo.xlsx.
    class MySumEngine : AbstractCalculationEngine
    {
        // The engine only handles the custom function "MYSUM"
        public override void Calculate(CalculationData data)
        {
            // Ensure we are processing the correct function (case‑insensitive)
            if (data.FunctionName.Equals("MYSUM", StringComparison.OrdinalIgnoreCase))
            {
                double sum = 0;

                // Iterate over all parameters passed to the function
                for (int i = 0; i < data.ParamCount; i++)
                {
                    // Get the parameter value; for simple cell references it is a ReferredArea
                    object param = data.GetParamValue(i);

                    // If the parameter is a ReferredArea, extract its single value
                    if (param is ReferredArea area)
                    {
                        // The area may contain a single cell; retrieve its value
                        object val = area.GetValue(0, 0);
                        sum += Convert.ToDouble(val);
                    }
                    else
                    {
                        // Direct literal numbers are returned as double or string
                        sum += Convert.ToDouble(param);
                    }
                }

                // Set the calculated result back to the engine
                data.CalculatedValue = sum;
            }
        }

        // No special recalculation logic required for this function
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

            // Populate sample data
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(25);

            // Set a formula that uses the custom function "MYSUM"
            sheet.Cells["B1"].Formula = "=MYSUM(A1, A2)";

            // Configure calculation options to use the custom engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new MySumEngine()
            };

            // Calculate formulas with the custom engine (lifecycle rule: calculate)
            workbook.CalculateFormula(options);

            // Output the result of the custom function
            Console.WriteLine("Result of MYSUM(A1, A2): " + sheet.Cells["B1"].Value);

            // Save the workbook (lifecycle rule: save)
            workbook.Save("CustomFunctionDemo.xlsx");
        }
    }
}
