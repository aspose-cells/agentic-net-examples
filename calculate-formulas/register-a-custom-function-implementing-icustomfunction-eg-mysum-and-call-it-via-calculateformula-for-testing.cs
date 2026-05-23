using System;
using Aspose.Cells;

namespace AsposeCellsCustomFunctionDemo
{
    // Custom calculation engine that implements a user‑defined function "MYSUM"
    class MySumEngine : AbstractCalculationEngine
    {
        // The engine only handles the custom function; other functions are processed by the default engine
        public override void Calculate(CalculationData data)
        {
            // Check if the current function is the custom one (case‑insensitive)
            if (data.FunctionName.Equals("MYSUM", StringComparison.OrdinalIgnoreCase))
            {
                // Ensure we have at least two parameters
                if (data.ParamCount >= 2)
                {
                    double sum = 0;

                    // Iterate over all parameters and add their numeric values
                    for (int i = 0; i < data.ParamCount; i++)
                    {
                        // Parameters are returned as objects; they may be ReferredArea or direct values
                        object param = data.GetParamValue(i);

                        // If the parameter is a ReferredArea (e.g., a cell reference), extract its value
                        if (param is ReferredArea area)
                        {
                            // Get the value from the first cell of the area
                            object val = area.GetValue(0, 0);
                            sum += Convert.ToDouble(val);
                        }
                        else
                        {
                            // Direct value (e.g., a constant number)
                            sum += Convert.ToDouble(param);
                        }
                    }

                    // Set the calculated result – this value will be returned to the worksheet cell
                    data.CalculatedValue = sum;
                }
                else
                {
                    // Not enough parameters – return Excel's #VALUE! error
                    data.CalculatedValue = "#VALUE!";
                }
            }
        }

        // No special force‑recalculation logic required for this demo
        public override bool ForceRecalculate(string functionName) => false;
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and obtain the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Populate cells with sample numeric data
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["B1"].PutValue(25);
            sheet.Cells["A2"].PutValue(5);
            sheet.Cells["B2"].PutValue(15);

            // 3. Set a formula that uses the custom function "MYSUM"
            //    The function will sum all supplied arguments (A1, B1, A2, B2)
            sheet.Cells["C1"].Formula = "=MYSUM(A1, B1, A2, B2)";

            // 4. Configure calculation options to use the custom engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new MySumEngine()
            };

            // 5. Perform calculation – the custom engine will be invoked for MYSUM
            workbook.CalculateFormula(options);

            // 6. Output the result of the custom function
            Console.WriteLine("Result of MYSUM(A1,B1,A2,B2): " + sheet.Cells["C1"].Value);

            // 7. Save the workbook (lifecycle rule: create → save)
            workbook.Save("CustomFunctionDemo.xlsx");
        }
    }
}