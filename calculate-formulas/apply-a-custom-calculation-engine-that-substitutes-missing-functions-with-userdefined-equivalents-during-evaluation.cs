// Title: Custom Calculation Engine in Aspose.Cells (.NET) for User‑Defined Functions
// Description: Shows how to create a workbook, populate cells, assign formulas that include the built‑in SUM and custom functions DOUBLE and CONCAT, implement a custom engine by extending AbstractCalculationEngine, set it via CalculationOptions.CustomEngine, calculate all formulas with wb.CalculateFormula, display the results, and save the workbook.
// Keywords: Aspose.Cells | custom calculation engine | AbstractCalculationEngine | user‑defined functions | C# | .NET | Excel formula extension | DOUBLE function | CONCAT function | wb.CalculateFormula | CalculationOptions | GitHub example
// Common Searches: Aspose.Cells custom calculation engine example | how to add user defined functions in Aspose.Cells | C# extend AbstractCalculationEngine | replace missing Excel functions with custom logic Aspose.Cells | calculate formulas with custom engine .NET
// Developer Intent: Create and register a custom calculation engine that implements undefined Excel functions and use it to evaluate workbook formulas in Aspose.Cells.
// Use Cases: Extend Aspose.Cells to support a DOUBLE function that returns twice the numeric argument. | Implement a CONCAT function that joins multiple cell values into a single string. | Run wb.CalculateFormula with a custom engine to combine built‑in and user‑defined functions, then export the results.
// AI Prompts: Generate a C# class inheriting from AbstractCalculationEngine that adds a POWER(base, exponent) function for Aspose.Cells. | Explain how to register a custom calculation engine in Aspose.Cells and evaluate formulas containing both native and custom functions. | Provide robust error‑handling code for parameter conversion inside a custom calculation engine method.

using System;
using Aspose.Cells;

namespace CustomEngineDemo
{
    // Shows how to create a workbook, populate cells, assign formulas that include the built‑in SUM and custom functions DOUBLE and CONCAT, implement a custom engine by extending AbstractCalculationEngine, set it via CalculationOptions.CustomEngine, calculate all formulas with wb.CalculateFormula, display the results, and save the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate cells with sample data
            ws.Cells["A1"].PutValue(5);
            ws.Cells["A2"].PutValue(10);
            ws.Cells["B1"].PutValue("Hello");
            ws.Cells["B2"].PutValue("World");

            // Built‑in function
            ws.Cells["C1"].Formula = "=SUM(A1:A2)";

            // Custom functions that are not built‑in
            ws.Cells["C2"].Formula = "=DOUBLE(A1)";
            ws.Cells["C3"].Formula = "=CONCAT(B1, \" \", B2)";

            // Set calculation options with the custom engine
            CalculationOptions opts = new CalculationOptions
            {
                CustomEngine = new MyCustomEngine()
            };

            // Calculate all formulas using the custom engine
            wb.CalculateFormula(opts);

            // Output results
            Console.WriteLine("SUM(A1:A2) = " + ws.Cells["C1"].Value);
            Console.WriteLine("DOUBLE(A1) = " + ws.Cells["C2"].Value);
            Console.WriteLine("CONCAT(B1, \" \", B2) = " + ws.Cells["C3"].StringValue);

            // Save the workbook
            wb.Save("CustomEngineResult.xlsx");
        }
    }

    // Custom calculation engine that provides implementations for missing functions
    class MyCustomEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            string func = data.FunctionName?.ToUpperInvariant();

            if (func == "DOUBLE")
            {
                // One numeric parameter, return its double
                object param = data.GetParamValue(0);
                double val = Convert.ToDouble(param);
                data.CalculatedValue = val * 2;
                return;
            }

            if (func == "CONCAT")
            {
                // Concatenate all parameters as strings
                var sb = new System.Text.StringBuilder();
                for (int i = 0; i < data.ParamCount; i++)
                {
                    object p = data.GetParamValue(i);
                    sb.Append(p?.ToString());
                }
                data.CalculatedValue = sb.ToString();
                return;
            }

            // For other functions, let the default engine handle them (do nothing)
        }

        // No need to force recalculation for these functions
        public override bool ForceRecalculate(string functionName) => false;
    }
}
