// Title: Implement a C# Aspose.Cells custom calculation engine that forces TODAY() to return a fixed date (2020‑01‑01)
// AI Prompts: Create a C# class inheriting AbstractCalculationEngine and override Calculate to assign new DateTime(2020,1,1) when the function name is TODAY. | Demonstrate how to set CalculationOptions.CustomEngine to the custom engine and invoke Workbook.CalculateFormula so that all TODAY() calls use the fixed date. | Write code that reads the value of a cell containing =TODAY(), prints the result, and saves the workbook to a file after using the custom engine.
// Common Searches: how to override TODAY() function in Aspose.Cells C# custom calculation engine | Aspose.Cells C# replace dynamic TODAY() with static date during formula evaluation | using AbstractCalculationEngine to return a constant date for TODAY in an Excel workbook | C# Aspose.Cells calculate formulas with a custom engine that fixes the TODAY value | save workbook after applying custom calculation engine for TODAY in Aspose.Cells
// Tags: Aspose.Cells custom AbstractCalculationEngine | override TODAY function Aspose.Cells | fixed date substitution Excel formula C# | CalculationOptions.CustomEngine usage | Workbook.CalculateFormula with custom engine

using System;
using Aspose.Cells;

namespace AsposeCellsCustomEngineDemo
{
    // Custom engine that replaces TODAY() with a fixed date (2020-01-01)
    // The example defines a FixedTodayEngine class that inherits AbstractCalculationEngine and overrides Calculate to set TODAY() to 2020‑01‑01. It configures CalculationOptions.CustomEngine with this engine, runs Workbook.CalculateFormula, prints the evaluated value from cell A1, and saves the workbook as FixedToday.xlsx.
    public class FixedTodayEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // Check if the function being calculated is TODAY
            if (data.FunctionName.Equals("TODAY", StringComparison.OrdinalIgnoreCase))
            {
                // Set the calculated value to the fixed date
                data.CalculatedValue = new DateTime(2020, 1, 1);
            }
            // For all other functions, do nothing and let the default engine handle them
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set a formula that uses TODAY()
            sheet.Cells["A1"].Formula = "=TODAY()";

            // Configure calculation options to use the custom engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new FixedTodayEngine()
            };

            // Calculate formulas with the custom engine
            workbook.CalculateFormula(options);

            // Output the result of the TODAY() formula
            Console.WriteLine("A1 value (fixed TODAY): " + sheet.Cells["A1"].Value);

            // Save the workbook
            workbook.Save("FixedToday.xlsx");
        }
    }
}
