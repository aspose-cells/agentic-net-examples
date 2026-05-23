using System;
using Aspose.Cells;

namespace CustomFunctionDemo
{
    // Custom calculation engine that implements a scalar function MYSUM
    // which sums all numeric values from the referenced range(s).
    public class MySumEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // Handle only the custom function named "MYSUM"
            if (data.FunctionName != null && data.FunctionName.Equals("MYSUM", StringComparison.OrdinalIgnoreCase))
            {
                double sum = 0.0;

                // Iterate through each parameter passed to the function
                for (int i = 0; i < data.ParamCount; i++)
                {
                    // Get the parameter value; for a range it will be a ReferredArea object
                    object param = data.GetParamValue(i);

                    if (param is ReferredArea area)
                    {
                        // Loop through all cells in the referred area and accumulate numeric values
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
                    else if (param is double d)
                    {
                        // Single numeric literal parameter
                        sum += d;
                    }
                    else
                    {
                        // Attempt conversion for other literal types
                        if (double.TryParse(Convert.ToString(param), out double parsed))
                        {
                            sum += parsed;
                        }
                    }
                }

                // Set the calculated scalar result
                data.CalculatedValue = sum;
            }
        }

        // No special force‑recalculation logic needed for this example
        public override bool ForceRecalculate(string functionName) => false;
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in a range A1:A3
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);

            // Set a formula that uses the custom scalar function MYSUM
            sheet.Cells["B1"].Formula = "=MYSUM(A1:A3)";

            // Configure calculation options to use the custom engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new MySumEngine()
            };

            // Perform calculation; the custom engine will be invoked
            workbook.CalculateFormula(options);

            // Retrieve and display the scalar result from cell B1
            Console.WriteLine("Result of MYSUM(A1:A3): " + sheet.Cells["B1"].Value);

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("CustomFunctionResult.xlsx");
        }
    }
}