using System;
using Aspose.Cells;

namespace CustomFunctionDemo
{
    // Custom function definition – specifies which parameters should be evaluated in array mode
    public class MyCustomFunctionDefinition : CustomFunctionDefinition
    {
        // For this demo, the first parameter (index 0) will be processed in array mode
        public override int[] GetArrayModeParameters(string functionName)
        {
            if (functionName.Equals("MYFUNC", StringComparison.OrdinalIgnoreCase))
            {
                return new int[] { 0 };
            }
            return base.GetArrayModeParameters(functionName);
        }
    }

    // Custom calculation engine – performs the actual calculation of the custom function
    public class MyCustomEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // Handle only the custom function we are interested in
            if (data.FunctionName.Equals("MYFUNC", StringComparison.OrdinalIgnoreCase))
            {
                // Retrieve the first parameter – because we marked it as array‑mode,
                // it may be returned as a 2‑D object array
                object param0 = data.GetParamValue(0);
                double sum = 0;

                if (param0 is object[,] arrayValues)
                {
                    // Iterate through all values in the array and accumulate the sum
                    foreach (object val in arrayValues)
                    {
                        if (val != null && double.TryParse(val.ToString(), out double d))
                        {
                            sum += d;
                        }
                    }
                }
                else
                {
                    // Single value case
                    if (param0 != null && double.TryParse(param0.ToString(), out double d))
                    {
                        sum = d;
                    }
                }

                // Second parameter is a normal scalar value
                object param1 = data.GetParamValue(1);
                if (param1 != null && double.TryParse(param1.ToString(), out double d2))
                {
                    sum += d2;
                }

                // Set the result – the engine must assign CalculatedValue
                data.CalculatedValue = sum;
            }
        }

        // Force recalculation for our custom function each time the workbook is calculated
        public override bool ForceRecalculate(string functionName)
        {
            return functionName.Equals("MYFUNC", StringComparison.OrdinalIgnoreCase);
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule – use provided creation method)
            Workbook wb = new Workbook();

            // Access the first worksheet
            Worksheet sheet = wb.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);
            sheet.Cells["B1"].PutValue(5); // scalar second argument

            // Update the workbook with our custom function definition
            wb.UpdateCustomFunctionDefinition(new MyCustomFunctionDefinition());

            // Set a formula that uses the custom function.
            // The first argument is a range (A1:A3) – it will be processed in array mode.
            // The second argument is a single cell (B1).
            sheet.Cells["C1"].Formula = "=MYFUNC(A1:A3, B1)";

            // Prepare calculation options with our custom engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new MyCustomEngine()
            };

            // Calculate all formulas using the custom engine
            wb.CalculateFormula(options);

            // Output the result of the custom function
            Console.WriteLine("Result of MYFUNC(A1:A3, B1): " + sheet.Cells["C1"].Value);

            // Save the workbook (lifecycle rule – use provided save method)
            wb.Save("CustomFunctionResult.xlsx");
        }
    }
}