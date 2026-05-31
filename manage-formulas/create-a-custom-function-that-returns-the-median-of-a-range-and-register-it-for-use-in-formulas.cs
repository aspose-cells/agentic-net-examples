using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Cells;

namespace CustomMedianFunctionDemo
{
    // Custom function definition to mark the first parameter (the range) as array‑mode
    public class MedianFunctionDefinition : CustomFunctionDefinition
    {
        public override int[] GetArrayModeParameters(string functionName)
        {
            // The function name we will use is "MYMEDIAN"
            if (functionName.Equals("MYMEDIAN", StringComparison.OrdinalIgnoreCase))
            {
                // Parameter index 0 (the first parameter) should be calculated in array mode
                return new int[] { 0 };
            }
            return null;
        }
    }

    // Custom calculation engine that implements the median logic
    public class MedianCalculationEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // We only handle our custom function
            if (!data.FunctionName.Equals("MYMEDIAN", StringComparison.OrdinalIgnoreCase))
                return; // Let the default engine handle other functions

            // Retrieve the first parameter value (already calculated in array mode)
            object param = data.GetParamValue(0);

            // Collect numeric values from the parameter
            List<double> values = new List<double>();

            if (param is object[,] array)
            {
                // Parameter is a 2‑D array (range)
                foreach (object item in array)
                {
                    if (item != null && double.TryParse(item.ToString(), out double d))
                        values.Add(d);
                }
            }
            else
            {
                // Single value case
                if (param != null && double.TryParse(param.ToString(), out double d))
                    values.Add(d);
            }

            // If there are no numeric values, result is #DIV/0! (or 0)
            if (values.Count == 0)
            {
                data.CalculatedValue = 0.0;
                return;
            }

            // Sort the values to compute median
            values.Sort();

            int n = values.Count;
            double median;
            if (n % 2 == 1)
            {
                // Odd count – middle element
                median = values[n / 2];
            }
            else
            {
                // Even count – average of two middle elements
                median = (values[(n / 2) - 1] + values[n / 2]) / 2.0;
            }

            data.CalculatedValue = median;
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data in A1:A6
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].PutValue(30);
            cells["A4"].PutValue(40);
            cells["A5"].PutValue(50);
            cells["A6"].PutValue(60);

            // Register the custom function definition (array‑mode for first parameter)
            workbook.UpdateCustomFunctionDefinition(new MedianFunctionDefinition());

            // Set a formula that uses the custom median function
            cells["B1"].Formula = "=MYMEDIAN(A1:A6)";

            // Prepare calculation options with our custom engine
            CalculationOptions calcOptions = new CalculationOptions
            {
                CustomEngine = new MedianCalculationEngine()
            };

            // Calculate all formulas using the custom engine
            workbook.CalculateFormula(calcOptions);

            // Output the result to console (optional)
            Console.WriteLine("Median of A1:A6 = " + cells["B1"].Value);

            // Save the workbook (create‑load‑save lifecycle)
            workbook.Save("MedianCustomFunctionDemo.xlsx");
        }
    }
}