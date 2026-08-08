// Title: Ignore #DIV/0! errors using CalculationOptions.IgnoreError in Aspose.Cells for .NET
// Description: Shows how to enable CalculationOptions.IgnoreError, evaluate a workbook with a division‑by‑zero formula, obtain the "#DIV/0!" string instead of an exception, compare the behavior when IgnoreError is false, and save the file.
// Keywords: Aspose.Cells | C# | CalculationOptions | IgnoreError | division by zero | #DIV/0! | formula error handling | Workbook.CalculateFormula | Excel error suppression | GitHub Aspose.Cells examples
// Common Searches: Aspose.Cells ignore division by zero | Set CalculationOptions.IgnoreError true C# | Prevent exception for #DIV/0! in Aspose.Cells | How to suppress formula errors Aspose.Cells | Calculate workbook without breaking on errors
// Developer Intent: Configure Aspose.Cells to ignore calculation errors so formulas that cause division‑by‑zero return the Excel error string rather than throwing an exception.
// Use Cases: Process large workbooks containing risky formulas without halting execution. | Preserve Excel error indicators (e.g., "#DIV/0!") in generated reports. | Implement custom error handling by comparing results with IgnoreError set to true versus false.
// AI Prompts: Generate C# code that sets CalculationOptions.IgnoreError = true and evaluates a workbook with a division‑by‑zero formula using Aspose.Cells. | Provide an example that catches the exception thrown when CalculationOptions.IgnoreError is false and logs the error message. | Explain how to read a cell's StringValue after calculation to detect the "#DIV/0!" error string.

using System;
using Aspose.Cells;

namespace AsposeCellsDivisionByZeroDemo
{
    // Shows how to enable CalculationOptions.IgnoreError, evaluate a workbook with a division‑by‑zero formula, obtain the "#DIV/0!" string instead of an exception, compare the behavior when IgnoreError is false, and save the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Set a formula that will cause division by zero
            cells["A1"].Formula = "=1/0";

            // Configure calculation options to ignore errors (including division‑by‑zero)
            CalculationOptions calcOptions = new CalculationOptions
            {
                IgnoreError = true   // Suppress errors during formula evaluation
            };

            // Perform calculation with the configured options
            workbook.CalculateFormula(calcOptions);

            // Output the result; Aspose.Cells returns the Excel error string "#DIV/0!"
            Console.WriteLine("Result of A1 after calculation with IgnoreError = true: " + cells["A1"].StringValue);

            // For comparison, calculate without ignoring errors to show exception handling
            try
            {
                // This will throw because IgnoreError defaults to true, but we set it explicitly to false here
                workbook.CalculateFormula(new CalculationOptions { IgnoreError = false });
            }
            catch (Exception ex)
            {
                Console.WriteLine("Calculation failed when IgnoreError = false: " + ex.Message);
            }

            // Save the workbook (optional)
            workbook.Save("DivisionByZeroDemo.xlsx");
        }
    }
}
