using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create a custom calculation engine that processes a user‑defined function DOUBLE()
        DoubleEngine customEngine = new DoubleEngine();

        // Set calculation options to use the custom engine
        CalculationOptions calcOptions = new CalculationOptions
        {
            CustomEngine = customEngine
        };

        // Calculate all formulas in the workbook using the custom engine
        workbook.CalculateFormula(calcOptions);

        // Save the workbook after calculation
        workbook.Save("output.xlsx");

        // Indicate whether the custom engine was invoked during calculation
        Console.WriteLine("Custom engine invoked: " + customEngine.Invoked);
    }
}

// Custom engine extending the default calculation engine
class DoubleEngine : AbstractCalculationEngine
{
    // Flag to show if the engine processed any DOUBLE() calls
    public bool Invoked { get; private set; }

    public override void Calculate(CalculationData data)
    {
        // Handle the custom function named DOUBLE
        if (data.FunctionName.Equals("DOUBLE", StringComparison.OrdinalIgnoreCase))
        {
            Invoked = true;

            // Expect at least one parameter; otherwise return an error value
            if (data.ParamCount > 0)
            {
                // Retrieve the first parameter's value
                object param = data.GetParamValue(0);
                double number = Convert.ToDouble(param);

                // Set the calculated result (double the input)
                data.CalculatedValue = number * 2;
            }
            else
            {
                data.CalculatedValue = "#VALUE!";
            }
        }
        // For all other functions, do nothing so the built‑in engine handles them
    }
}