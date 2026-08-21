// Title: Register a Custom COMPOUNDINTEREST Function in Aspose.Cells (.NET)
// Description: Defines a CompoundInterestEngine that inherits AbstractCalculationEngine, intercepts the COMPOUNDINTEREST call with three arguments (principal, rate, periods), computes principal × ((1 + rate)^periods − 1), returns the result or "#VALUE!" for invalid input, attaches the engine via CalculationOptions.CustomEngine, applies the formula =COMPOUNDINTEREST(A1,A2,A3) in a worksheet, triggers calculation, prints the value, and saves the workbook as an XLSX file.
// Keywords: Aspose.Cells | custom function | COMPOUNDINTEREST | AbstractCalculationEngine | C# | .NET | financial formula | compound interest calculation | CalculationOptions | Excel custom function | workbook calculation
// Common Searches: how to add a custom function in Aspose.Cells | register custom calculation engine Aspose.Cells .NET | compound interest formula C# Aspose.Cells example | use AbstractCalculationEngine for financial calculations | custom Excel functions with Aspose.Cells
// Developer Intent: Create and register a COMPOUNDINTEREST custom function in Aspose.Cells to compute compound interest directly within workbook formulas.
// Use Cases: Embed a reusable compound‑interest calculator in any Excel‑like workbook generated with Aspose.Cells. | Build financial models that automatically update when principal, rate, or period cells change. | Produce investment or loan reports where interest values are derived without external scripts.
// AI Prompts: Write C# code that implements a COMPOUNDINTEREST custom function for Aspose.Cells using AbstractCalculationEngine. | Show how to handle missing or extra arguments for a custom financial function in Aspose.Cells. | Explain the steps to configure CalculationOptions.CustomEngine so all workbook formulas use the custom engine.

using System;
using Aspose.Cells;

namespace AsposeCellsCompoundInterestDemo
{
    // Custom calculation engine that implements the COMPOUNDINTEREST function
    // Defines a CompoundInterestEngine that inherits AbstractCalculationEngine, intercepts the COMPOUNDINTEREST call with three arguments (principal, rate, periods), computes principal × ((1 + rate)^periods − 1), returns the result or "#VALUE!" for invalid input, attaches the engine via CalculationOptions.CustomEngine, applies the formula =COMPOUNDINTEREST(A1,A2,A3) in a worksheet, triggers calculation, prints the value, and saves the workbook as an XLSX file.
    public class CompoundInterestEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // Check if the function being evaluated is our custom function
            if (data.FunctionName.Equals("COMPOUNDINTEREST", StringComparison.OrdinalIgnoreCase))
            {
                // Expecting three parameters: principal, rate, periods
                if (data.ParamCount == 3)
                {
                    // Retrieve parameter values (they are returned as objects, usually double)
                    double principal = Convert.ToDouble(data.GetParamValue(0));
                    double rate = Convert.ToDouble(data.GetParamValue(1));
                    double periods = Convert.ToDouble(data.GetParamValue(2));

                    // Compound interest formula: principal * ((1 + rate) ^ periods - 1)
                    double result = principal * (Math.Pow(1 + rate, periods) - 1);

                    // Set the calculated value so Aspose.Cells can use it
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
            // Create a new workbook (lifecycle rule: create)
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate input values
            cells["A1"].PutValue(1000);      // Principal
            cells["A2"].PutValue(0.05);      // Rate (5%)
            cells["A3"].PutValue(10);        // Periods (years)

            // Set the formula that uses the custom function
            cells["B1"].Formula = "=COMPOUNDINTEREST(A1, A2, A3)";

            // Configure calculation options to use our custom engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new CompoundInterestEngine()
            };

            // Calculate all formulas in the workbook using the custom engine
            wb.CalculateFormula(options);

            // Output the result
            Console.WriteLine("Compound Interest Result: " + cells["B1"].Value);

            // Save the workbook (lifecycle rule: save)
            wb.Save("CompoundInterestResult.xlsx", SaveFormat.Xlsx);
        }
    }
}
