using System;
using Aspose.Cells;

namespace AsposeCellsCustomFunctionDemo
{
    // Custom calculation engine that implements a user‑defined function "MYSUM"
    public class MySumEngine : AbstractCalculationEngine
    {
        // This engine does not need array‑mode or literal parameters
        public override void Calculate(CalculationData data)
        {
            // Handle only the custom function "MYSUM"
            if (data.FunctionName.Equals("MYSUM", StringComparison.OrdinalIgnoreCase))
            {
                double sum = 0;

                // Iterate over all parameters passed to the function
                for (int i = 0; i < data.ParamCount; i++)
                {
                    // Parameters are returned as ReferredArea objects
                    var area = (ReferredArea)data.GetParamValue(i);

                    // Get the value from the first cell of the area (single‑cell arguments)
                    object valObj = area.GetValue(0, 0);
                    if (valObj != null && double.TryParse(valObj.ToString(), out double val))
                    {
                        sum += val;
                    }
                }

                // Set the calculated result – this value will appear in the cell containing the formula
                data.CalculatedValue = sum;
            }
        }

        // Force recalculation of the custom function each time the workbook is calculated
        public override bool ForceRecalculate(string functionName)
        {
            return functionName.Equals("MYSUM", StringComparison.OrdinalIgnoreCase);
        }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Populate sample data
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(25);

            // 3. Set a formula that uses the custom function "MYSUM"
            sheet.Cells["C1"].Formula = "=MYSUM(A1, A2)";

            // 4. Configure calculation options to use the custom engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new MySumEngine()
            };

            // 5. Calculate all formulas in the workbook using the custom engine
            workbook.CalculateFormula(options);

            // 6. Output the result of the custom function
            Console.WriteLine("Result of MYSUM(A1, A2): " + sheet.Cells["C1"].Value);

            // 7. Save the workbook (demonstrates that the custom function result is persisted)
            workbook.Save("CustomFunctionResult.xlsx");
        }
    }
}