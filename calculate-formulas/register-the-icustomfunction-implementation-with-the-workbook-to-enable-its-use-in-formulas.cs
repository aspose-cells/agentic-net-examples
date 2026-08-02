using System;
using Aspose.Cells;

namespace AsposeCellsCustomFunctionDemo
{
    // Custom calculation engine that implements a user‑defined function named MYFUNC
    public class MyCustomFunctionEngine : AbstractCalculationEngine
    {
        // This method is called for every function encountered during calculation
        public override void Calculate(CalculationData data)
        {
            // Check if the function name matches our custom function (case‑insensitive)
            if (string.Equals(data.FunctionName, "MYFUNC", StringComparison.OrdinalIgnoreCase))
            {
                // Expecting exactly two parameters; retrieve their values
                // GetParamValue returns the evaluated value of the parameter
                object param0 = data.GetParamValue(0);
                object param1 = data.GetParamValue(1);

                // Convert parameters to double (handle possible nulls)
                double val0 = param0 != null ? Convert.ToDouble(param0) : 0.0;
                double val1 = param1 != null ? Convert.ToDouble(param1) : 0.0;

                // Example logic: return the sum of the two parameters multiplied by 2
                data.CalculatedValue = (val0 + val1) * 2;
            }
            // For any other function, do nothing – the default engine will handle it
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle rule: use provided creation method)
            Workbook wb = new Workbook();

            // Access the first worksheet
            Worksheet sheet = wb.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data that the custom function will use
            cells["A1"].PutValue(5);
            cells["A2"].PutValue(7);

            // Set a formula that uses the custom function MYFUNC
            cells["A3"].Formula = "=MYFUNC(A1, A2)";

            // Configure calculation options to use our custom engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new MyCustomFunctionEngine()
            };

            // Calculate all formulas in the workbook using the custom engine
            wb.CalculateFormula(options);

            // Output the result of the custom function
            Console.WriteLine("Result of MYFUNC(A1, A2): " + cells["A3"].Value);

            // Save the workbook (lifecycle rule: use provided save method)
            wb.Save("CustomFunctionDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}