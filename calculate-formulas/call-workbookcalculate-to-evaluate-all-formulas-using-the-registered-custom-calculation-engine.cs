using System;
using Aspose.Cells;

namespace AsposeCellsCustomEngineDemo
{
    // Custom calculation engine that implements a user‑defined function CUSTOMSUM
    public class CustomEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // Handle only the custom function name (case‑insensitive)
            if (data.FunctionName.Equals("CUSTOMSUM", StringComparison.OrdinalIgnoreCase))
            {
                double sum = 0;

                // Iterate through all parameters passed to the function
                for (int i = 0; i < data.ParamCount; i++)
                {
                    object param = data.GetParamValue(i);

                    // The parameter may be a numeric value (double, int, etc.)
                    if (param is double d)
                        sum += d;
                    else if (param is int iVal)
                        sum += iVal;
                    // Add more type checks if needed (e.g., decimal, long)
                }

                // Set the calculated result – this will be written back to the cell
                data.CalculatedValue = sum;
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // -------------------- Create workbook --------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);

            // Use the custom function in a formula
            sheet.Cells["B1"].Formula = "=CUSTOMSUM(A1,A2)";

            // -------------------- Set calculation options --------------------
            CalculationOptions options = new CalculationOptions
            {
                // Register the custom engine so that CalculateFormula uses it
                CustomEngine = new CustomEngine(),
                // Optional: keep default behavior for other settings
                IgnoreError = true,
                Recursive = true
            };

            // -------------------- Calculate all formulas --------------------
            // This evaluates every formula in the workbook using the custom engine
            workbook.CalculateFormula(options);

            // -------------------- Output result --------------------
            Console.WriteLine("Result of CUSTOMSUM(A1,A2) in B1: " + sheet.Cells["B1"].Value);

            // -------------------- Save workbook --------------------
            workbook.Save("CustomEngineResult.xlsx");
        }
    }
}