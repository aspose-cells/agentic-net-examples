using System;
using Aspose.Cells;

namespace AsposeCellsCustomFunctionDemo
{
    // Custom calculation engine that implements a scalar numeric function "REFSUM"
    // The function receives one or more reference parameters and returns the sum of all referenced cells.
    public class RefSumEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // Ensure we are handling the expected custom function (case‑insensitive)
            if (!string.Equals(data.FunctionName, "REFSUM", StringComparison.OrdinalIgnoreCase))
                return;

            double total = 0.0;

            // Iterate over all parameters supplied to the function
            for (int i = 0; i < data.ParamCount; i++)
            {
                // Get the parameter value; for a reference it will be a ReferredArea object
                object param = data.GetParamValue(i);

                if (param is ReferredArea area)
                {
                    // Loop through the rectangular area and accumulate numeric values
                    for (int r = area.StartRow; r <= area.EndRow; r++)
                    {
                        for (int c = area.StartColumn; c <= area.EndColumn; c++)
                        {
                            object cellVal = area.GetValue(r, c);
                            if (cellVal != null && double.TryParse(cellVal.ToString(), out double d))
                            {
                                total += d;
                            }
                        }
                    }
                }
                else
                {
                    // Parameter might be a plain scalar value (e.g., a constant number)
                    if (param != null && double.TryParse(param.ToString(), out double d))
                    {
                        total += d;
                    }
                }
            }

            // Set the calculated scalar result – this value will be returned to the worksheet cell
            data.CalculatedValue = total;
        }

        // No special handling for shared formulas; default implementation is sufficient
        public override bool ForceRecalculate(string functionName) => false;
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and obtain the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Populate sample data that will be referenced by the custom function
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);
            sheet.Cells["B1"].PutValue(5);   // additional reference to demonstrate multiple parameters
            sheet.Cells["B2"].PutValue(15);

            // 3. Set a formula that uses the custom scalar function REF​SUM
            //    The function receives two parameters: a range A1:A3 and a range B1:B2
            sheet.Cells["C1"].Formula = "=REFSUM(A1:A3, B1:B2)";

            // 4. Configure calculation options to use the custom engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new RefSumEngine()
            };

            // 5. Perform calculation – the custom engine will be invoked for REF​SUM
            workbook.CalculateFormula(options);

            // 6. Retrieve and display the scalar result
            Console.WriteLine("Result of REF​SUM(A1:A3, B1:B2): " + sheet.Cells["C1"].Value);

            // 7. Save the workbook (optional, demonstrates lifecycle compliance)
            workbook.Save("RefSumDemo.xlsx");
        }
    }
}