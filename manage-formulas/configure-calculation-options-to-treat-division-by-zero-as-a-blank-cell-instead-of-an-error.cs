// Title: Aspose.Cells .NET – Convert Division‑by‑Zero Errors to Blank Cells
// Description: Demonstrates how to configure Aspose.Cells CalculationOptions to ignore errors, run Workbook.CalculateFormula, detect #DIV/0! results, replace them with an empty string, and save the workbook.
// Keywords: Aspose.Cells division by zero | Aspose.Cells ignore error | CalculationOptions IgnoreError | C# replace #DIV/0! with blank | .NET Excel formula error handling
// Common Searches: Aspose.Cells hide #DIV/0! error | C# Aspose.Cells treat division by zero as empty cell | Aspose.Cells CalculationOptions ignore errors example | How to clear error values after CalculateFormula in Aspose.Cells | Replace Excel formula errors with blank using Aspose.Cells .NET
// Developer Intent: Replace a division‑by‑zero error in an Aspose.Cells workbook with an empty cell instead of #DIV/0!.
// Use Cases: Financial reports where invalid ratios should appear blank for readability. | Automated Excel generation that must hide calculation errors before distribution. | Templates that automatically clear error cells to prevent confusing end‑users.
// AI Prompts: Provide C# code that sets CalculationOptions.IgnoreError, runs CalculateFormula, and clears #DIV/0! cells in Aspose.Cells. | Show a method to scan a worksheet after calculation and replace any error value with an empty string using Aspose.Cells. | Explain alternative ways to suppress division‑by‑zero errors in Aspose.Cells without manual string checks.

using System;
using Aspose.Cells;

namespace AsposeCellsDivisionByZeroDemo
{
    // Demonstrates how to configure Aspose.Cells CalculationOptions to ignore errors, run Workbook.CalculateFormula, detect #DIV/0! results, replace them with an empty string, and save the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Set a formula that causes division by zero
            cells["A1"].Formula = "=10/0";

            // Configure calculation options:
            // - IgnoreError = true suppresses exceptions during calculation.
            // - The result of the error will still be an error value, so we replace it with a blank.
            CalculationOptions calcOptions = new CalculationOptions
            {
                IgnoreError = true
            };

            // Perform calculation with the configured options
            workbook.CalculateFormula(calcOptions);

            // After calculation, check if the cell contains a division‑by‑zero error
            // (error values are represented as strings starting with "#").
            if (cells["A1"].StringValue.StartsWith("#"))
            {
                // Treat the error as a blank cell
                cells["A1"].PutValue(string.Empty);
            }

            // Save the workbook to verify the result
            workbook.Save("DivisionByZeroHandled.xlsx");
        }
    }
}
