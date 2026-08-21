// Title: Evaluate an Excel formula string with Worksheet.CalculateFormula and custom CalculationOptions in Aspose.Cells for .NET
// Description: Shows how to build a Workbook, populate cells, define a formula as a string, configure a CalculationOptions object, and invoke Worksheet.CalculateFormula(formula, options) to obtain the result without inserting the formula into any worksheet cell.
// Keywords: Worksheet.CalculateFormula | Aspose.Cells | CalculationOptions | evaluate formula string | C# Excel formula evaluation | formula without cell | custom calculation options | programmatic formula calculation | .NET Aspose.Cells API | Excel engine evaluation
// Common Searches: Aspose.Cells evaluate formula string without cell | Worksheet.CalculateFormula with CalculationOptions C# | How to use CalculationOptions in Aspose.Cells | Programmatically calculate Excel formula in .NET | Aspose.Cells formula engine custom options
// Developer Intent: Compute the result of an Excel formula supplied as a string using Worksheet.CalculateFormula, optionally applying custom CalculationOptions, without writing the formula to the worksheet.
// Use Cases: Generate totals from a range for a report while keeping the workbook unchanged | Validate user‑entered formulas on a server before saving them | Run performance‑critical calculations with specific settings such as disabling recursion or ignoring errors
// AI Prompts: Create C# code that evaluates a nested IF formula using Worksheet.CalculateFormula with CalculationOptions.Recursive = true. | Show how to capture and handle errors from Worksheet.CalculateFormula when CalculationOptions.IgnoreError is set to false. | Demonstrate using a custom calculation engine via CalculationOptions while evaluating a formula string in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaEvaluation
{
    // Demonstrates evaluating a formula string directly using Worksheet.CalculateFormula
    // with a custom CalculationOptions instance, without writing the formula to a cell.
    // Shows how to build a Workbook, populate cells, define a formula as a string, configure a CalculationOptions object, and invoke Worksheet.CalculateFormula(formula, options) to obtain the result without inserting the formula into any worksheet cell.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data that the formula can reference.
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);

            // Define the formula to evaluate. It references the cells populated above.
            string formula = "=SUM(A1:A3)";

            // Create a CalculationOptions instance.
            // Here we use default options, but you could customize properties such as
            // IgnoreError, Recursive, CustomEngine, etc.
            CalculationOptions calcOptions = new CalculationOptions
            {
                // Example of customizing an option:
                // IgnoreError = false,
                // Recursive = true
            };

            // Evaluate the formula directly without placing it in any cell.
            object result = sheet.CalculateFormula(formula, calcOptions);

            // Output the result.
            Console.WriteLine($"Result of formula \"{formula}\": {result}");
        }
    }
}
