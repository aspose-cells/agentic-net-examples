using System;
using Aspose.Cells;

namespace CustomEngineDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate cells that will be used as parameters for the custom function
            ws.Cells["A1"].PutValue(5);
            ws.Cells["A2"].PutValue(7);

            // Set a formula that calls the custom function MYFUNC
            ws.Cells["A3"].Formula = "=MYFUNC(A1, A2)";

            // Create calculation options and assign the custom engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new MyCustomEngine()
            };

            // Register the custom engine and calculate all formulas in the workbook
            wb.CalculateFormula(options);

            // Display the result of the custom function
            Console.WriteLine("Result of MYFUNC(A1, A2): " + ws.Cells["A3"].Value);

            // Save the workbook (optional, demonstrates that the result is persisted)
            wb.Save("CustomEngineResult.xlsx");
        }
    }

    // Custom calculation engine that extends the default Aspose.Cells engine
    public class MyCustomEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // Process only the custom function named "MYFUNC"
            if (data.FunctionName.Equals("MYFUNC", StringComparison.OrdinalIgnoreCase))
            {
                // Retrieve the first parameter as a ReferredArea and extract its value
                ReferredArea area1 = (ReferredArea)data.GetParamValue(0);
                double val1 = Convert.ToDouble(area1.GetValue(0, 0));

                // Retrieve the second parameter as a ReferredArea and extract its value
                ReferredArea area2 = (ReferredArea)data.GetParamValue(1);
                double val2 = Convert.ToDouble(area2.GetValue(0, 0));

                // Example custom logic: return the product of the two parameters
                data.CalculatedValue = val1 * val2;
            }
        }
    }
}