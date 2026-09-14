// Title: How to suppress #DIV/0! errors by setting CalculationOptions.IgnoreError to true when calculating formulas in Aspose.Cells (C#)
// AI Prompts: Write C# code that creates a workbook, inserts a division‑by‑zero formula, enables CalculationOptions.IgnoreError, and runs CalculateFormula. | Generate a C# example that first calculates with CalculationOptions.IgnoreError = true, then repeats the calculation with IgnoreError = false and catches the resulting exception. | Explain step‑by‑step how to configure Aspose.Cells CalculationOptions to ignore formula errors and retrieve the cell value after calculation.
// Common Searches: Aspose.Cells C# calculate workbook formulas while ignoring division by zero errors | set CalculationOptions.IgnoreError to true in Aspose.Cells example | how to handle #DIV/0! in Aspose.Cells CalculateFormula method | C# Aspose.Cells ignore formula errors during calculation
// Tags: Aspose.Cells CalculationOptions.IgnoreError | C# ignore #DIV/0! Aspose.Cells | Aspose.Cells calculate formulas with error suppression | handling division by zero in Aspose.Cells workbook | Aspose.Cells formula evaluation error handling C#

using System;
using Aspose.Cells;

// Demonstrates configuring CalculationOptions.IgnoreError to true to suppress division‑by‑zero (#DIV/0!) errors when calculating formulas in an Aspose.Cells workbook, and shows the exception thrown when the option is disabled.
class DivisionByZeroIgnoreErrorDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Set a normal value and a formula that causes division by zero
        worksheet.Cells["A1"].PutValue(10);
        worksheet.Cells["B1"].Formula = "=A1/0"; // Division by zero

        // Configure calculation options to ignore errors
        CalculationOptions calcOptions = new CalculationOptions
        {
            IgnoreError = true // Suppress errors such as #DIV/0!
        };

        // Calculate all formulas using the options
        workbook.CalculateFormula(calcOptions);

        // Display the result of the division‑by‑zero formula
        Console.WriteLine("B1 value after calculation (ignore error): " + worksheet.Cells["B1"].StringValue);

        // Demonstrate behavior when errors are not ignored
        try
        {
            workbook.CalculateFormula(new CalculationOptions { IgnoreError = false });
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error when not ignoring: " + ex.Message);
        }
    }
}
