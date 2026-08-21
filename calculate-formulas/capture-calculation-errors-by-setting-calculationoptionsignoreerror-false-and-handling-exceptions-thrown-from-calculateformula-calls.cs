// Title: Capture Aspose.Cells formula errors in C# by setting CalculationOptions.IgnoreError = false
// Description: Shows how to insert an invalid Excel formula, disable error suppression with CalculationOptions.IgnoreError, call Workbook.CalculateFormula inside a try‑catch to catch the exception, then re‑run the calculation with errors ignored and read the resulting cell value.
// Keywords: Aspose.Cells | C# | CalculateFormula | CalculationOptions | IgnoreError | formula error handling | exception handling | Excel formula validation | workbook.CalculateFormula | error capture
// Common Searches: Aspose.Cells catch calculation exception | CalculateFormula throws error when IgnoreError false | how to handle invalid Excel formula Aspose.Cells .NET | disable error ignoring in Aspose.Cells calculation | retrieve error message from workbook.CalculateFormula
// Developer Intent: Detect and manage formula calculation failures by turning off error suppression in Aspose.Cells.
// Use Cases: Validate all formulas in an uploaded workbook and abort processing on the first error. | Provide a strict calculation mode for end‑users that surfaces the exact cause of a formula failure. | Switch between tolerant and strict calculation at runtime based on user preferences or configuration.
// AI Prompts: Write C# code that logs the full stack trace when Workbook.CalculateFormula throws an exception with CalculationOptions.IgnoreError set to false. | Explain how to toggle CalculationOptions.IgnoreError in a .NET microservice that processes Excel files using Aspose.Cells. | Create a retry pattern that fixes a formula after catching the exception, then recalculates the workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsCalculationErrorDemo
{
    // Shows how to insert an invalid Excel formula, disable error suppression with CalculationOptions.IgnoreError, call Workbook.CalculateFormula inside a try‑catch to catch the exception, then re‑run the calculation with errors ignored and read the resulting cell value.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Set a formula that will cause a calculation error (non‑existent function)
            cells["A1"].Formula = "=NONEXISTENTFUNC(1,2)";

            // Prepare calculation options with IgnoreError set to false
            CalculationOptions options = new CalculationOptions
            {
                IgnoreError = false   // Do not ignore errors – an exception will be thrown
            };

            // Attempt to calculate formulas and capture any errors
            try
            {
                workbook.CalculateFormula(options);
                Console.WriteLine("Calculation completed without errors (unexpected).");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Caught calculation error as expected: " + ex.Message);
            }

            // Now calculate again, this time ignoring errors
            options.IgnoreError = true;   // Suppress errors
            try
            {
                workbook.CalculateFormula(options);
                Console.WriteLine("Calculation completed while ignoring errors.");
            }
            catch (Exception ex)
            {
                // This block should not be reached when IgnoreError = true
                Console.WriteLine("Unexpected error: " + ex.Message);
            }

            // Optionally, display the resulting value in the cell (will be the error string)
            Console.WriteLine("Cell A1 value after ignoring errors: " + cells["A1"].StringValue);
        }
    }
}
