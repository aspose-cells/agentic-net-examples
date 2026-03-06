using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create a custom calculation engine that handles the HYPERLINK function
        var customEngine = new HyperlinkEngine();

        // Set calculation options to use the custom engine
        CalculationOptions calcOptions = new CalculationOptions
        {
            CustomEngine = customEngine
        };

        // Perform formula calculation with the custom engine
        workbook.CalculateFormula(calcOptions);

        // Output whether the custom engine was invoked
        Console.WriteLine("Custom engine invoked: " + customEngine.Invoked);

        // Save the workbook after calculation
        workbook.Save("output.xlsx");
    }

    // Custom engine derived from AbstractCalculationEngine
    class HyperlinkEngine : AbstractCalculationEngine
    {
        // Flag to indicate if the engine processed a HYPERLINK function
        public bool Invoked { get; private set; }

        public override void Calculate(CalculationData data)
        {
            // Check for the HYPERLINK function (case‑insensitive)
            if (data.FunctionName.Equals("HYPERLINK", StringComparison.OrdinalIgnoreCase))
            {
                Invoked = true;

                // Retrieve the display text (second parameter) as literal text
                string displayText = data.GetParamText(1);

                // Set a custom calculated value
                data.CalculatedValue = "[Custom Processed] " + displayText;
            }
            // For all other functions, let the default engine handle them
        }
    }
}