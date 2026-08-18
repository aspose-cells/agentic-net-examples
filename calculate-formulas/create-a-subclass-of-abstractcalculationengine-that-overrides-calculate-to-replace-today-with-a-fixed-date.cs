// Title: C# Custom Aspose.Cells Calculation Engine that Returns a Fixed Date for TODAY()
// Description: Shows how to inherit from Aspose.Cells.AbstractCalculationEngine, override the Calculate method to detect the TODAY function and supply a predefined DateTime (e.g., 2020‑01‑01), while letting the default engine handle all other formulas. Includes setup with CalculationOptions and workbook formula evaluation.
// Keywords: Aspose.Cells | AbstractCalculationEngine | custom calculation engine | override TODAY | fixed date | C# Excel formula | CalculateFormula | CalculationOptions | Excel serial date | reproducible tests
// Common Searches: Aspose.Cells replace TODAY with constant date | C# custom AbstractCalculationEngine example | How to fix TODAY() value in Aspose.Cells | CalculateFormula with custom engine C# | Set static date for Excel TODAY function using Aspose
// Developer Intent: Replace the dynamic TODAY() function with a static date during formula evaluation.
// Use Cases: Create deterministic test workbooks where TODAY() must not change between runs. | Generate audit‑ready financial reports that require a fixed reference date. | Run scenario analyses that need a consistent “current” date across multiple calculations.
// AI Prompts: Write a C# class that inherits from Aspose.Cells.AbstractCalculationEngine and returns January 1 2020 for any TODAY() call, delegating other functions to the default engine. | Demonstrate how to configure CalculationOptions to use a custom engine that overrides TODAY() with a fixed date and then calculate all formulas in a workbook. | Explain why assigning a DateTime to CalculationData.CalculatedValue is automatically converted to an Excel serial number by Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsCustomEngineDemo
{
    // Custom calculation engine that overrides the TODAY function
    // Shows how to inherit from Aspose.Cells.AbstractCalculationEngine, override the Calculate method to detect the TODAY function and supply a predefined DateTime (e.g., 2020‑01‑01), while letting the default engine handle all other formulas. Includes setup with CalculationOptions and workbook formula evaluation.
    public class FixedTodayEngine : AbstractCalculationEngine
    {
        // Fixed date to be returned for TODAY()
        private readonly DateTime _fixedDate = new DateTime(2020, 1, 1);

        public override void Calculate(CalculationData data)
        {
            // Check if the function being evaluated is TODAY (case‑insensitive)
            if (string.Equals(data.FunctionName, "TODAY", StringComparison.OrdinalIgnoreCase))
            {
                // Set the calculated value to the fixed date.
                // Aspose.Cells expects a double representing an Excel serial date,
                // but assigning a DateTime object works as the library converts it.
                data.CalculatedValue = _fixedDate;
                // No need to call base.Calculate() because the method is abstract.
            }
            // For all other functions we let the default engine handle the calculation
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: use provided creation pattern)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Put some sample data (optional)
            sheet.Cells["A1"].PutValue(123);

            // Set a formula that uses the TODAY function
            sheet.Cells["B1"].Formula = "=TODAY()";

            // Configure calculation options to use the custom engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new FixedTodayEngine()
            };

            // Calculate all formulas in the workbook using the custom engine
            workbook.CalculateFormula(options);

            // Output the result to verify the TODAY function was replaced
            Console.WriteLine("B1 value (fixed TODAY): " + sheet.Cells["B1"].Value);

            // Save the workbook (lifecycle rule: use provided save pattern)
            workbook.Save("FixedTodayDemo.xlsx");
        }
    }
}
