// Title: Create a custom scalar function MYSUM in Aspose.Cells for .NET to sum a referenced range and return a numeric result
// AI Prompts: Write a C# class that inherits AbstractCalculationEngine and implements a scalar function named MYSUM which iterates over a ReferredArea to add numeric and parsable string values. | Show how to attach the custom MySumEngine to CalculationOptions and invoke =MYSUM(A1:A3) in a worksheet to obtain a single numeric output. | Extend the MySumEngine example to ignore non‑numeric cells and ensure the function returns a double without triggering a full workbook recalculation.
// Common Searches: aspnet custom function MYSUM Aspose.Cells sum range | how to implement scalar custom calculation engine in Aspose.Cells C# | example of using ReferredArea to aggregate values in Aspose.Cells formula | register custom engine with CalculationOptions to calculate custom formulas | return double result from user‑defined function in Aspose.Cells workbook
// Tags: custom scalar function Aspose.Cells | AbstractCalculationEngine sum range | MYSUM custom formula .NET | CalculationOptions register custom engine | ReferredArea numeric aggregation | handle string numbers in Aspose.Cells function

using System;
using Aspose.Cells;

namespace CustomFunctionDemo
{
    // Custom calculation engine that implements a scalar function "MYSUM"
    // which sums all numeric values in the referenced range.
    // The example defines MySumEngine inheriting AbstractCalculationEngine, overrides Calculate to sum numeric and parsable string values from a ReferredArea passed to the custom function MYSUM, assigns the sum to CalculatedValue, registers the engine via CalculationOptions, applies the formula =MYSUM(A1:A3) on a worksheet, and outputs the scalar numeric result.
    public class MySumEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // Handle only the custom function "MYSUM"
            if (data.FunctionName.Equals("MYSUM", StringComparison.OrdinalIgnoreCase))
            {
                double sum = 0.0;

                // Expect exactly one parameter (a range reference)
                if (data.ParamCount > 0)
                {
                    // Get the parameter value; for a range it will be a ReferredArea object
                    object param = data.GetParamValue(0);
                    if (param is ReferredArea area)
                    {
                        // Iterate through all cells in the area and accumulate numeric values
                        for (int r = area.StartRow; r <= area.EndRow; r++)
                        {
                            for (int c = area.StartColumn; c <= area.EndColumn; c++)
                            {
                                object cellValue = area.GetValue(r, c);
                                if (cellValue is double d)
                                {
                                    sum += d;
                                }
                                else
                                {
                                    // Try to convert other types (e.g., string numbers) to double
                                    if (double.TryParse(Convert.ToString(cellValue), out double parsed))
                                    {
                                        sum += parsed;
                                    }
                                }
                            }
                        }
                    }
                }

                // Set the scalar result back to the engine
                data.CalculatedValue = sum;
            }
        }

        // No need to force recalculation for this simple function
        public override bool ForceRecalculate(string functionName) => false;
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in A1:A3
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);

            // Set a formula that uses the custom function MYSUM
            sheet.Cells["B1"].Formula = "=MYSUM(A1:A3)";

            // Configure calculation options to use the custom engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new MySumEngine()
            };

            // Perform calculation; the custom engine will be invoked
            workbook.CalculateFormula(options);

            // Retrieve and display the scalar numeric result
            Console.WriteLine("Result of MYSUM(A1:A3): " + sheet.Cells["B1"].Value);

            // Save the workbook (optional, demonstrates lifecycle compliance)
            workbook.Save("CustomFunctionResult.xlsx");
        }
    }
}
