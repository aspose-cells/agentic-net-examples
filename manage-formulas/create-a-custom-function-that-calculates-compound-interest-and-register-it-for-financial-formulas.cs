// Title: Register a custom COMPOUNDINTEREST function in Aspose.Cells (.NET) for compound‑interest calculations
// Description: This example shows how to create a CompoundInterestEngine that inherits AbstractCalculationEngine, implements the COMPOUNDINTEREST function (principal, rate, periods), registers it via CalculationOptions.CustomEngine, applies the formula =COMPOUNDINTEREST(A1,B1,C1) in a worksheet, calculates the workbook, outputs the result, and saves the file.
// Keywords: Aspose.Cells custom function | C# compound interest | AbstractCalculationEngine | CalculationOptions.CustomEngine | financial formulas Aspose | Excel custom formula .NET | register custom engine Aspose.Cells
// Common Searches: how to add a custom function in Aspose.Cells C# | register custom calculation engine for financial formulas | compound interest formula Aspose.Cells example | use CalculationOptions.CustomEngine to extend Excel functions | Aspose.Cells custom financial function tutorial
// Developer Intent: Create and register a COMPOUNDINTEREST function so Aspose.Cells can evaluate it like a native Excel formula.
// Use Cases: Add a bespoke future‑value calculation to budgeting spreadsheets. | Implement depreciation or amortization logic in asset‑management workbooks. | Provide tax or payroll calculations that are not covered by built‑in Excel functions.
// AI Prompts: Generate a C# class that implements a custom Aspose.Cells engine for a Net Present Value (NPV) function. | Show how to register multiple custom functions in a single CalculationOptions.CustomEngine instance. | Write unit tests for CompoundInterestEngine covering correct results and argument‑error handling.

using System;
using Aspose.Cells;

namespace CustomCompoundInterestDemo
{
    // Custom calculation engine that implements the COMPOUNDINTEREST function
    // This example shows how to create a CompoundInterestEngine that inherits AbstractCalculationEngine, implements the COMPOUNDINTEREST function (principal, rate, periods), registers it via CalculationOptions.CustomEngine, applies the formula =COMPOUNDINTEREST(A1,B1,C1) in a worksheet, calculates the workbook, outputs the result, and saves the file.
    public class CompoundInterestEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // Check if the function being calculated is our custom function
            if (data.FunctionName.Equals("COMPOUNDINTEREST", StringComparison.OrdinalIgnoreCase))
            {
                // Expect three parameters: principal, rate, periods
                if (data.ParamCount == 3)
                {
                    // Retrieve parameter values (they are already evaluated by the engine)
                    double principal = Convert.ToDouble(data.GetParamValue(0));
                    double rate = Convert.ToDouble(data.GetParamValue(1));
                    double periods = Convert.ToDouble(data.GetParamValue(2));

                    // Compound interest formula: A = P * (1 + r) ^ n
                    double result = principal * Math.Pow(1 + rate, periods);

                    // Set the calculated value so Aspose.Cells returns it
                    data.CalculatedValue = result;
                }
                else
                {
                    // Incorrect number of arguments – return an error value
                    data.CalculatedValue = "#VALUE!";
                }
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            Cells cells = sheet.Cells;

            // Input data for compound interest calculation
            // Principal amount
            cells["A1"].PutValue(1000);
            // Annual interest rate (e.g., 5% => 0.05)
            cells["B1"].PutValue(0.05);
            // Number of periods (years)
            cells["C1"].PutValue(10);

            // Set the formula that uses the custom function
            cells["D1"].Formula = "=COMPOUNDINTEREST(A1, B1, C1)";

            // Prepare calculation options with our custom engine
            CalculationOptions opts = new CalculationOptions
            {
                CustomEngine = new CompoundInterestEngine()
            };

            // Calculate all formulas in the workbook using the custom engine
            wb.CalculateFormula(opts);

            // Output the result of the compound interest calculation
            Console.WriteLine("Compound Interest Result (Cell D1): " + cells["D1"].Value);

            // Save the workbook (optional, demonstrates lifecycle usage)
            wb.Save("CompoundInterestResult.xlsx", SaveFormat.Xlsx);
        }
    }
}
