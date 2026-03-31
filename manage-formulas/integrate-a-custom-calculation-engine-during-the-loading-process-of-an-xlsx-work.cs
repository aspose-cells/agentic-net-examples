using System;
using Aspose.Cells;

namespace CustomEngineDemo
{
    class Program
    {
        static void Main()
        {
            // Load the workbook with LoadOptions.
            // ParsingFormulaOnOpen = false loads formulas as raw strings (no parsing at load time).
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.ParsingFormulaOnOpen = false;

            // Path to the source XLSX file.
            string inputPath = "input.xlsx";
            Workbook wb = new Workbook(inputPath, loadOptions);

            // Create an instance of the custom calculation engine.
            var customEngine = new MyCustomEngine();

            // Set calculation options to use the custom engine.
            CalculationOptions calcOptions = new CalculationOptions
            {
                CustomEngine = customEngine,
                // Continue calculation even if some formulas are unsupported.
                IgnoreError = true
            };

            // Calculate all formulas in the workbook using the custom engine.
            wb.CalculateFormula(calcOptions);

            // Save the workbook after calculation.
            string outputPath = "output.xlsx";
            wb.Save(outputPath, SaveFormat.Xlsx);

            // Indicate whether the custom engine was invoked.
            Console.WriteLine("Custom engine invoked: " + customEngine.Invoked);
        }
    }

    // Custom calculation engine that processes a user‑defined function MYFUNC.
    public class MyCustomEngine : AbstractCalculationEngine
    {
        public bool Invoked { get; private set; }

        public override void Calculate(CalculationData data)
        {
            // Handle the custom function MYFUNC.
            if (data.FunctionName.Equals("MYFUNC", StringComparison.OrdinalIgnoreCase))
            {
                // Assume MYFUNC takes two numeric parameters and returns their product.
                double param1 = Convert.ToDouble(data.GetParamValue(0));
                double param2 = Convert.ToDouble(data.GetParamValue(1));
                data.CalculatedValue = param1 * param2;
                Invoked = true;
            }
            // For all other functions let the default engine handle the calculation.
        }
    }
}