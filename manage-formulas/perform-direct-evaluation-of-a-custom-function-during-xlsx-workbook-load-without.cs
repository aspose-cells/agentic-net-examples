using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load options – we can control formula parsing on open if needed
        LoadOptions loadOptions = new LoadOptions
        {
            // Skip parsing formulas while opening (optional, demonstrates load control)
            ParsingFormulaOnOpen = false
        };

        // Load an existing XLSX workbook (replace with your actual file path)
        Workbook workbook = new Workbook("template.xlsx", loadOptions);

        // Prepare calculation options with a custom engine that knows how to evaluate MYFUNC
        CalculationOptions calcOptions = new CalculationOptions
        {
            CustomEngine = new MyCustomEngine()
        };

        // Directly evaluate the custom function without placing it in any worksheet cell
        // The formula is evaluated as if it were in cell A1 of the first worksheet
        object result = workbook.Worksheets[0].CalculateFormula("=MYFUNC(7,3)", calcOptions);

        Console.WriteLine($"Result of MYFUNC(7,3): {result}");

        // Optionally save the workbook (no changes were made to cells)
        workbook.Save("Result.xlsx");
    }

    // Custom calculation engine that processes the user‑defined function MYFUNC
    class MyCustomEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // Handle only the custom function we are interested in
            if (string.Equals(data.FunctionName, "MYFUNC", StringComparison.OrdinalIgnoreCase))
            {
                // Retrieve the two parameters passed to MYFUNC
                double param0 = Convert.ToDouble(data.GetParamValue(0));
                double param1 = Convert.ToDouble(data.GetParamValue(1));

                // Example custom logic: multiply the two parameters
                data.CalculatedValue = param0 * param1;
            }
        }
    }
}