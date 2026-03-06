using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Create a custom calculation engine that knows how to evaluate MYFUNC
        var customEngine = new MyCustomEngine();

        // Set calculation options to use the custom engine
        CalculationOptions options = new CalculationOptions
        {
            CustomEngine = customEngine
        };

        // Directly calculate the custom function without writing it to any cell
        string formula = "=MYFUNC(5, 7)";
        object result = worksheet.CalculateFormula(formula, options);

        Console.WriteLine($"Result of {formula} = {result}");

        // Optionally save the workbook (no changes were made to cells)
        workbook.Save("output.xlsx");
    }

    // Custom engine that implements the logic for MYFUNC
    class MyCustomEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // Handle only the custom function named MYFUNC (case‑insensitive)
            if (data.FunctionName.Equals("MYFUNC", StringComparison.OrdinalIgnoreCase))
            {
                double sum = 0;

                // Iterate over all parameters passed to the function
                for (int i = 0; i < data.ParamCount; i++)
                {
                    object param = data.GetParamValue(i);

                    // Parameters may be plain numbers (double) or a ReferredArea (range)
                    if (param is double d)
                    {
                        sum += d;
                    }
                    else if (param is ReferredArea area)
                    {
                        // Take the first cell value from the area for simplicity
                        object val = area.GetValue(0, 0);
                        if (val is double dv)
                            sum += dv;
                        else if (double.TryParse(val?.ToString(), out double parsed))
                            sum += parsed;
                    }
                }

                // Example custom logic: return double the sum of the parameters
                data.CalculatedValue = sum * 2;
            }
        }
    }
}