using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Set up calculation options with a custom engine
        CalculationOptions options = new CalculationOptions
        {
            CustomEngine = new MyCustomEngine()
        };

        // Directly calculate a custom function without placing it in a cell
        string formula = "=MYFUNC(5, 10)";
        object result = worksheet.CalculateFormula(formula, options);

        Console.WriteLine("Result of MYFUNC(5,10): " + result);

        // Optionally store the result in a cell and save the workbook
        worksheet.Cells["A1"].PutValue(result);
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }

    // Custom calculation engine that processes the MYFUNC function
    class MyCustomEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // Handle only the custom function named MYFUNC
            if (string.Equals(data.FunctionName, "MYFUNC", StringComparison.OrdinalIgnoreCase))
            {
                double sum = 0;

                // Iterate through all parameters and sum numeric values
                for (int i = 0; i < data.ParamCount; i++)
                {
                    object param = data.GetParamValue(i);
                    if (param is double d)
                        sum += d;
                    else if (param is int iVal)
                        sum += iVal;
                }

                // Example custom logic: return double the sum
                data.CalculatedValue = sum * 2;
            }
        }
    }
}