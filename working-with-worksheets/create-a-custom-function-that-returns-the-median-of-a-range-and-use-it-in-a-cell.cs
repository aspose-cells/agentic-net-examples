using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsCustomMedian
{
    // Custom function definition – tells the parser which parameters need array‑mode calculation
    class MedianFunctionDefinition : CustomFunctionDefinition
    {
        // The first (and only) parameter is a range, so it must be evaluated in array mode
        public override int[] GetArrayModeParameters(string functionName)
        {
            if (functionName.Equals("MEDIAN_CUSTOM", StringComparison.OrdinalIgnoreCase))
                return new int[] { 0 };   // index of the range parameter
            return null;
        }
    }

    // Custom calculation engine – implements the actual median logic
    class MedianEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            if (!data.FunctionName.Equals("MEDIAN_CUSTOM", StringComparison.OrdinalIgnoreCase))
                return; // let default engine handle other functions

            // The parameter is supplied in array mode → object[,] (2‑D array)
            object param = data.GetParamValue(0);
            if (param is object[,] array)
            {
                // Flatten the 2‑D array into a list of doubles, ignoring nulls
                List<double> values = new List<double>();
                int rows = array.GetLength(0);
                int cols = array.GetLength(1);
                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        if (array[r, c] != null && double.TryParse(array[r, c].ToString(), out double v))
                            values.Add(v);
                    }
                }

                if (values.Count == 0)
                {
                    data.CalculatedValue = null; // empty range → #VALUE!
                    return;
                }

                // Sort and compute median
                values.Sort();
                int n = values.Count;
                double median = (n % 2 == 1)
                    ? values[n / 2]
                    : (values[(n / 2) - 1] + values[n / 2]) / 2.0;

                data.CalculatedValue = median;
            }
            else
            {
                // If not an array (should not happen for our definition), fallback to default handling
                data.CalculatedValue = null;
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // 2. Populate a sample range A1:A5
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].PutValue(30);
            cells["A4"].PutValue(40);
            cells["A5"].PutValue(50);

            // 3. Define the custom function name and set the formula in B1
            string formula = "=MEDIAN_CUSTOM(A1:A5)";
            FormulaParseOptions parseOptions = new FormulaParseOptions
            {
                // Attach our custom function definition so the parser knows the range is array‑mode
                CustomFunctionDefinition = new MedianFunctionDefinition(),
                Parse = true
            };
            cells["B1"].SetFormula(formula, parseOptions);

            // 4. Prepare calculation options with our custom engine
            CalculationOptions calcOptions = new CalculationOptions
            {
                CustomEngine = new MedianEngine(),
                // For performance we can disable recursive recalculation if not needed
                Recursive = false
            };

            // 5. Calculate all formulas using the custom engine
            wb.CalculateFormula(calcOptions);

            // 6. Output the result of the custom median function
            Console.WriteLine("Median of A1:A5 (custom function) = " + cells["B1"].Value);

            // 7. Save the workbook (lifecycle rule – use the standard Save method)
            wb.Save("CustomMedianDemo.xlsx");
        }
    }
}