// Title: Use Worksheet.CalculateFormula with a custom CalculationOptions to evaluate an Excel formula string without inserting it into the worksheet (C#)
// AI Prompts: Write C# code that creates a Workbook, sets a cell value, configures CalculationOptions to ignore errors, and calls the appropriate method to compute a formula string directly. | Show how to evaluate an Excel formula on the fly using Aspose.Cells by passing a formula and a CalculationOptions instance to the calculation API.
// Common Searches: asp.net evaluate Excel formula string using Aspose.Cells without adding to worksheet | how to apply CalculationOptions when calling Worksheet.CalculateFormula in C# | ignore errors during formula calculation with Aspose.Cells CalculationOptions example | run ad‑hoc formula evaluation in Aspose.Cells C# code sample
// Tags: custom CalculationOptions for formula evaluation | ad‑hoc formula evaluation Aspose.Cells | ignore errors during Aspose.Cells calculation | process Excel formula string in C# | calculate formula without worksheet insertion

using Aspose.Cells;
using System;

// The example creates a new Workbook, writes a value to cell A1, sets CalculationOptions to ignore errors, and uses Worksheet.CalculateFormula to evaluate the formula "=SUM(1,2,3)*A1" directly, outputting the result without adding the formula to the worksheet.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set up calculation options
            CalculationOptions calcOptions = new CalculationOptions
            {
                // Ignore errors during calculation
                IgnoreError = true
            };

            // Example: put a value in a cell that the formula will reference
            sheet.Cells["A1"].PutValue(10);

            // Formula string to evaluate (note the leading '=')
            string formula = "=SUM(1, 2, 3) * A1";

            // Evaluate the formula without adding it to the worksheet
            object result = sheet.CalculateFormula(formula, calcOptions);

            // Display the result
            Console.WriteLine("Result: " + result);
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
