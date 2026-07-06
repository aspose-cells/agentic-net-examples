using System;
using Aspose.Cells;

namespace AsposeCellsCustomEngineDemo
{
    // Custom calculation engine that replaces the old ICustomFunction implementation.
    // It inherits from AbstractCalculationEngine and overrides the required members.
    public class MyCustomEngine : AbstractCalculationEngine
    {
        // Indicates whether this engine needs the literal text of parameters.
        // For this simple example we don't need it, so keep the default (false).
        public override bool IsParamLiteralRequired => false;

        // Indicates whether this engine needs parameters to be calculated in array mode.
        // Not required for our simple sum function.
        public override bool IsParamArrayModeRequired => false;

        // If you want built‑in functions to be processed by this engine, set this to true.
        // Here we only handle our custom function, so return false.
        public override bool ProcessBuiltInFunctions => false;

        // Force recalculation for volatile custom functions.
        // Our function is deterministic, so we return false.
        public override bool ForceRecalculate(string functionName) => false;

        // Core calculation logic. Aspose.Cells will call this for each function occurrence.
        public override void Calculate(CalculationData data)
        {
            // Handle only the custom function named "MYFUNC".
            if (string.Equals(data.FunctionName, "MYFUNC", StringComparison.OrdinalIgnoreCase))
            {
                // Ensure we have the expected number of parameters.
                if (data.ParamCount >= 2)
                {
                    // Retrieve the first parameter value.
                    // GetParamValue returns an object; for cell references it is a ReferredArea.
                    object param0 = data.GetParamValue(0);
                    object param1 = data.GetParamValue(1);

                    double val0 = ConvertToDouble(param0);
                    double val1 = ConvertToDouble(param1);

                    // Set the calculated result; this value will be written back to the cell.
                    data.CalculatedValue = val0 + val1;
                }
                else
                {
                    // Not enough parameters – return an error value.
                    data.CalculatedValue = "#VALUE!";
                }
            }
            // For any other function let the default engine handle it (do nothing here).
        }

        // Helper to convert a parameter (which may be a ReferredArea or a direct value) to double.
        private double ConvertToDouble(object param)
        {
            if (param is ReferredArea area)
            {
                // For a cell reference we take the value from the first cell of the area.
                object cellValue = area.GetValue(0, 0);
                return Convert.ToDouble(cellValue);
            }
            // Direct numeric value.
            return Convert.ToDouble(param);
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data.
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);

            // Use the custom function in a formula.
            sheet.Cells["A3"].Formula = "=MYFUNC(A1, A2)";

            // Set calculation options to use the custom engine.
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new MyCustomEngine(),
                IgnoreError = false,
                Recursive = true
            };

            // Perform calculation with the custom engine.
            workbook.CalculateFormula(options);

            // Output the result.
            Console.WriteLine("Result of MYFUNC(A1, A2): " + sheet.Cells["A3"].Value);

            // Save the workbook in both Excel and PDF formats.
            workbook.Save("CustomEngineResult.xlsx");
            workbook.Save("CustomEngineResult.pdf");
        }
    }
}