using System;
using Aspose.Cells;

namespace AsposeCellsCustomArrayFunctionDemo
{
    // Custom function definition specifying which parameters need array‑mode calculation
    class MyFuncDefinition : CustomFunctionDefinition
    {
        // The first (and only) parameter of MYFUNC should be evaluated in array mode
        public override int[] GetArrayModeParameters(string functionName)
        {
            if (functionName.Equals("MYFUNC", StringComparison.OrdinalIgnoreCase))
                return new int[] { 0 };
            return null;
        }
    }

    // Custom calculation engine that implements the logic of MYFUNC
    class MyFuncEngine : AbstractCalculationEngine
    {
        // The engine requires array‑mode values for its parameters
        public override bool IsParamArrayModeRequired => true;

        public override void Calculate(CalculationData data)
        {
            if (!data.FunctionName.Equals("MYFUNC", StringComparison.OrdinalIgnoreCase))
                return; // Let Aspose.Cells handle other functions

            // Obtain the parameter as an array (rows x columns)
            // Use generous limits; actual size will be trimmed automatically
            object[][] paramArray = data.GetParamValueInArrayMode(0, 0, 0);

            double sum = 0;
            foreach (object[] row in paramArray)
            {
                foreach (object item in row)
                {
                    if (item != null && double.TryParse(item.ToString(), out double val))
                        sum += val;
                }
            }

            // Return the sum as the function result
            data.CalculatedValue = sum;
        }

        // No special forced recalculation logic needed
        public override bool ForceRecalculate(string functionName) => false;
    }

    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Create a workbook and populate sample data
            // ------------------------------------------------------------
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // Column B values
            cells["B1"].PutValue(1);
            cells["B2"].PutValue(2);
            cells["B3"].PutValue(3);

            // Column C values
            cells["C1"].PutValue(10);
            cells["C2"].PutValue(20);
            cells["C3"].PutValue(30);

            // ------------------------------------------------------------
            // 2. Register custom function definition
            // ------------------------------------------------------------
            wb.UpdateCustomFunctionDefinition(new MyFuncDefinition());

            // ------------------------------------------------------------
            // 3. Set a formula that uses the custom function with an array argument
            // ------------------------------------------------------------
            // The expression B1:B3+C1:C3 creates an array of sums (1+10, 2+20, 3+30)
            // MYFUNC will receive this array and sum its elements (i.e., 63)
            cells["D1"].Formula = "=MYFUNC(B1:B3+C1:C3)";

            // ------------------------------------------------------------
            // 4. Prepare calculation options with the custom engine
            // ------------------------------------------------------------
            CalculationOptions calcOpts = new CalculationOptions
            {
                CustomEngine = new MyFuncEngine()
            };

            // ------------------------------------------------------------
            // 5. Calculate the workbook (the custom function will be invoked)
            // ------------------------------------------------------------
            wb.CalculateFormula(calcOpts);

            Console.WriteLine("Result of MYFUNC in cell D1: " + cells["D1"].Value); // Expected 63

            // ------------------------------------------------------------
            // 6. Demonstrate CalculateArrayFormula to obtain the raw array result
            // ------------------------------------------------------------
            string arrayFormula = "=MYFUNC(B1:B3+C1:C3)";
            object[][] arrayResult = ws.CalculateArrayFormula(arrayFormula, calcOpts);

            // The result is a 1x1 array because MYFUNC returns a single scalar value
            Console.WriteLine("CalculateArrayFormula returned: " + arrayResult[0][0]);

            // ------------------------------------------------------------
            // 7. Save the workbook (lifecycle rule compliance)
            // ------------------------------------------------------------
            wb.Save("CustomArrayFunctionDemo.xlsx");
        }
    }
}