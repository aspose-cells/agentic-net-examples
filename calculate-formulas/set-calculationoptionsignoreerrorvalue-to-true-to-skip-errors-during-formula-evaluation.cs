// Title: Skip formula errors in Aspose.Cells .NET with CalculationOptions.IgnoreError
// Description: Shows how to enable CalculationOptions.IgnoreError so workbook.CalculateFormula suppresses errors (e.g., #DIV/0!), returns a clean cell value, and saves the workbook.
// Keywords: Aspose.Cells | CalculationOptions | IgnoreError | C# .NET | skip formula errors | suppress #DIV/0 | calculate formulas | workbook.CalculateFormula | error handling in Excel | Excel automation
// Common Searches: Aspose.Cells ignore formula errors | Set CalculationOptions.IgnoreError true | Suppress #DIV/0 when calculating workbook | Calculate formulas without exceptions Aspose.Cells | C# skip errors during Excel calculation
// Developer Intent: Evaluate all formulas in a workbook while preventing error values or exceptions from interrupting the calculation process.
// Use Cases: Generate reports where some cells contain placeholder formulas that may be invalid until data is filled. | Batch‑process spreadsheets and continue calculations even if individual cells produce errors such as division by zero. | Automate template calculations and save the file without exposing #DIV/0 or #VALUE! errors to end users.
// AI Prompts: How do I configure Aspose.Cells CalculationOptions to ignore errors and retrieve the resulting cell values? | Provide a C# example that runs workbook.CalculateFormula with IgnoreError enabled and shows how to identify cells that originally had errors. | Explain how to log cells that generated errors when CalculationOptions.IgnoreError is set to true in Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to enable CalculationOptions.IgnoreError so workbook.CalculateFormula suppresses errors (e.g., #DIV/0!), returns a clean cell value, and saves the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook (creation rule)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Set a formula that will generate an error (division by zero)
        sheet.Cells["A1"].Formula = "=1/0";

        // Configure calculation options to ignore errors
        CalculationOptions options = new CalculationOptions
        {
            IgnoreError = true // Skip errors during formula evaluation
        };

        // Calculate all formulas using the specified options
        workbook.CalculateFormula(options);

        // Display the result; with IgnoreError=true the error is suppressed
        Console.WriteLine("A1 value after calculation with IgnoreError=true: " + sheet.Cells["A1"].StringValue);

        // Save the workbook (saving rule)
        workbook.Save("IgnoreErrorDemo.xlsx");
    }
}
