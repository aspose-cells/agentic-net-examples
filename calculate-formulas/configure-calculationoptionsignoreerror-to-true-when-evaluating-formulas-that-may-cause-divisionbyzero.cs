// Title: Ignore Division‑by‑Zero Errors in Aspose.Cells for .NET Using CalculationOptions.IgnoreError
// Description: Shows how to enable CalculationOptions.IgnoreError so that a formula like =10/0 returns an error string (e.g., #DIV/0!) instead of throwing an exception, and demonstrates the opposite behavior when the flag is disabled.
// Keywords: Aspose.Cells | C# | .NET | CalculationOptions | IgnoreError | division by zero | formula error handling | workbook.CalculateFormula | error string | Excel error handling
// Common Searches: Aspose.Cells ignore division by zero | CalculationOptions.IgnoreError true example | C# Aspose.Cells prevent exception on invalid formula | How to get #DIV/0! instead of exception Aspose.Cells | Set calculation options to ignore errors Aspose.Cells
// Developer Intent: Configure calculation options so that formulas causing errors return an error indicator rather than raising an exception.
// Use Cases: Process spreadsheets containing invalid formulas without halting the calculation pipeline. | Display standard Excel error symbols (e.g., #DIV/0!) in generated reports instead of aborting execution. | Run bulk formula evaluations across many workbooks while programmatically detecting cells that contain error strings.
// AI Prompts: Provide C# code that calculates all formulas in a workbook with Aspose.Cells while ignoring any errors, then logs cells that contain error strings. | Show how to toggle CalculationOptions.IgnoreError between true and false to compare results for a division‑by‑zero formula. | Explain how Aspose.Cells represents ignored errors in cell values and how to detect them programmatically.

using System;
using Aspose.Cells;

// Shows how to enable CalculationOptions.IgnoreError so that a formula like =10/0 returns an error string (e.g., #DIV/0!) instead of throwing an exception, and demonstrates the opposite behavior when the flag is disabled.
class DivisionByZeroIgnoreErrorDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Set a formula that causes division by zero
        cells["A1"].Formula = "=10/0";

        // Configure calculation options to ignore errors
        CalculationOptions options = new CalculationOptions
        {
            IgnoreError = true
        };

        // Calculate formulas using the options
        workbook.CalculateFormula(options);

        // Display the result; when errors are ignored the cell shows the error string
        Console.WriteLine("A1 result with IgnoreError=true: " + cells["A1"].StringValue);

        // Demonstrate behavior when errors are not ignored (exception is thrown)
        try
        {
            workbook.CalculateFormula(new CalculationOptions { IgnoreError = false });
        }
        catch (Exception ex)
        {
            Console.WriteLine("Calculation threw exception as expected: " + ex.Message);
        }

        // Save the workbook (optional)
        workbook.Save("DivisionByZeroDemo.xlsx");
    }
}
