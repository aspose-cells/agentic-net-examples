// Title: How to configure Aspose.Cells CalculationOptions to ignore division‑by‑zero errors in C#
// Common Searches: Aspose.Cells ignore #DIV/0 error during workbook.CalculateFormula in C# | set CalculationOptions.IgnoreError true to suppress division by zero in Aspose.Cells | C# evaluate Excel formulas with Aspose.Cells without raising DIV/0 exceptions | configure Aspose.Cells calculation options to return zero for invalid formulas | using CalculationOptions to handle division by zero in Aspose.Cells .NET
// Tags: CalculationOptions.IgnoreError property | suppress division by zero Aspose.Cells | Aspose.Cells formula error handling | C# workbook.CalculateFormula with error suppression | Excel formula evaluation with error ignoring

using System;
using Aspose.Cells;

// The example creates a workbook, inserts a formula that divides by zero, sets CalculationOptions.IgnoreError to true so the error is ignored and the result becomes zero, evaluates all formulas with these options, outputs the result, and saves the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Insert a formula that may cause division by zero (e.g., =A1/0)
        // A1 is left empty, so the denominator will be zero.
        Cell formulaCell = sheet.Cells["B1"];
        formulaCell.Formula = "=A1/0";

        // Configure calculation options to ignore errors such as division‑by‑zero
        CalculationOptions calcOptions = new CalculationOptions();
        calcOptions.IgnoreError = true; // When true, errors are ignored and result is 0

        // Evaluate all formulas using the configured options
        workbook.CalculateFormula(calcOptions);

        // Optionally, read the evaluated result
        Console.WriteLine("Result of B1: " + formulaCell.Value);

        // Save the workbook (lifecycle rule: save)
        workbook.Save("ResultWorkbook.xlsx");
    }
}
